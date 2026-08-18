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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.path.join(BASE_DIR, 'forecast_cache.json')
EXCEL_DEFAULT = os.path.join(BASE_DIR, 'historico.xlsx')

app = Flask(__name__)
CORS(app)

DATA_CACHE_IN_MEMORY = None

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
    return jsonify({"error": "ALERTA CRÍTICA: No se encontró index.html"}), 404

@app.route('/favicon.ico')
def favicon():
    return '', 204

def clean_num(val, default=0.0):
    if pd.isna(val) or val is None: return default
    try:
        val_str = str(val).strip().replace(',', '.')
        val_str = re.sub(r'[^0-9.]', '', val_str)
        return float(val_str) if val_str else default
    except Exception:
        return default

def parse_aht_to_seconds(val):
    if pd.isna(val) or val is None: return 180.0
    secs = 180.0
    if isinstance(val, (int, float)):
        secs = float(val)
    elif hasattr(val, 'hour') and hasattr(val, 'minute') and hasattr(val, 'second'):
        secs = val.hour * 3600 + val.minute * 60 + val.second
    else:
        val_str = str(val).strip()
        if ':' in val_str:
            parts = val_str.split(':')
            try:
                if len(parts) == 3: secs = int(parts[0]) * 3600 + int(parts[1]) * 60 + float(parts[2])
                elif len(parts) == 2: secs = int(parts[0]) * 3600 + float(parts[1])
            except: pass
        else:
            try: secs = float(val_str)
            except: pass
    if 0 < secs <= 15: secs = secs * 60.0
    return secs if secs > 0 else 180.0

def format_aht_str(seconds):
    if pd.isna(seconds) or seconds is None or seconds <= 0: return "00:00"
    secs = int(round(seconds))
    return f"{secs // 3600:02d}:{(secs % 3600) // 60:02d}"

def calcular_agentes_requeridos_erlang_c_rapido(A, target_sl=80.0):
    if A <= 0: return 0
    z = 0.84 if target_sl >= 80 else 0.5
    return int(math.ceil(A + (z * math.sqrt(A)) + 0.5))

def erlang_c_sl_optimizado(A, N, AHT, target_time):
    if N <= A or A <= 0 or N <= 0: return 0.0
    try:
        intensity = N - A
        pw = max(0.01, min(1.0, A / N))
        sl = 1.0 - (pw * math.exp(-intensity * (target_time / AHT)))
        return round(max(0.0, min(100.0, sl * 100.0)), 1)
    except: return 0.0

def parse_time_str(t_str):
    if not t_str: return None
    t = str(t_str).lower().replace('hrs', '').replace(' ', '')
    is_pm = 'pm' in t
    is_am = 'am' in t
    t = t.replace('am', '').replace('pm', '')
    t = re.sub(r'[^\d:]', '', t)
    if not t: return None
    if ':' not in t: t += ':00'
    try:
        parts = t.split(':')
        hh, mm = int(parts[0]), int(parts[1])
        if is_pm and hh < 12: hh += 12
        if is_am and hh == 12: hh = 0
        return hh * 60 + mm
    except: return None

def esta_en_ventana_servicio(campana, intervalo_str):
    camp_key = str(campana).strip().lower()
    minutos_inter = parse_time_str(intervalo_str)
    if minutos_inter is None: return True
    if 'liverpool' in camp_key or 'suburbia' in camp_key:
        return (9 * 60) <= minutos_inter < (21 * 60)
    return True

def get_col_safe(df, keywords, default_idx=0):
    for kw in keywords:
        for c in df.columns:
            if kw.lower() in str(c).strip().lower():
                return c
    return df.columns[default_idx] if len(df.columns) > default_idx else df.columns[0]

def extraer_datos_plantilla_rapido(xls_file):
    try:
        sheet_plantilla = None
        for s in xls_file.sheet_names:
            if 'plat' in s.lower() or 'plan' in s.lower() or 'rost' in s.lower():
                sheet_plantilla = s
                break
        if not sheet_plantilla: return {}, {}

        df_p = pd.read_excel(xls_file, sheet_name=sheet_plantilla, engine='openpyxl')
        dias_cols = ['lunes', 'martes', 'miércoles', 'miercoles', 'jueves', 'viernes', 'sábado', 'sabado', 'domingo']
        
        col_id = get_col_safe(df_p, ['id', 'agente', 'empleado'], 0)
        col_camp = get_col_safe(df_p, ['campaña', 'campana', 'skill', 'servicio'], 2 if len(df_p.columns) >= 3 else 0)

        df_p['Camp_Clean'] = df_p[col_camp].astype(str).str.strip().str.title()
        
        hc_total_campana = df_p.groupby('Camp_Clean')[col_id].nunique().to_dict()
        hc_total_campana['General'] = int(df_p[col_id].nunique())

        hc_diario_campana = {}

        for _, row in df_p.iterrows():
            camp_norm = str(row['Camp_Clean'])
            for col_name in df_p.columns:
                dia_key = str(col_name).strip().lower()
                if any(d in dia_key for d in dias_cols):
                    base_day = next(d for d in dias_cols if d in dia_key)
                    if base_day == 'miercoles': base_day = 'miércoles'
                    if base_day == 'sabado': base_day = 'sábado'

                    val_turno = str(row[col_name]).strip().lower()
                    if val_turno and not any(x in val_turno for x in ['dd', 'descanso', 'falta', 'vacacion', 'baja', 'nan']):
                        k_dia = (camp_norm, base_day)
                        k_gen = ('General', base_day)
                        hc_diario_campana[k_dia] = hc_diario_campana.get(k_dia, 0) + 1
                        hc_diario_campana[k_gen] = hc_diario_campana.get(k_gen, 0) + 1

        return hc_total_campana, hc_diario_campana
    except Exception as e:
        print("Error analizando hoja plantilla:", e)
        return {}, {}

def holt_winters_fit_predict_rapido(series, n_preds=180, season_len=7, alpha=0.2, beta=0.05, gamma=0.2):
    n = len(series)
    if n < season_len * 2:
        return [float(np.mean(series)) if len(series) > 0 else 100.0] * n_preds
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

def procesar_archivo_excel_rapido(file_source, target_sl=80.0, target_time=20.0, merma=0.20, dias_futuros=180):
    global DATA_CACHE_IN_MEMORY
    
    if DATA_CACHE_IN_MEMORY is not None and isinstance(DATA_CACHE_IN_MEMORY, list) and len(DATA_CACHE_IN_MEMORY) > 0:
        return DATA_CACHE_IN_MEMORY

    xls_file = pd.ExcelFile(file_source, engine='openpyxl')
    
    hc_total_map, hc_diario_map = extraer_datos_plantilla_rapido(xls_file)

    sheet_calls = xls_file.sheet_names[0]
    for s in xls_file.sheet_names:
        if 'llam' in s.lower() or 'hist' in s.lower() or 'datos' in s.lower():
            sheet_calls = s
            break

    df_raw = pd.read_excel(xls_file, sheet_name=sheet_calls, engine='openpyxl')

    col_calls = get_col_safe(df_raw, ['recibida', 'llamada', 'call', 'volumen', 'ofrecida'], 5)
    col_aht = get_col_safe(df_raw, ['aht', 'tmo', 'duracio', 'handle'], 12 if len(df_raw.columns) > 12 else 0)
    col_camp = get_col_safe(df_raw, ['campaña', 'campana', 'skill', 'servicio', 'ring group'], 3 if len(df_raw.columns) > 3 else 0)
    col_inter = get_col_safe(df_raw, ['intervalo', 'hora', 'time'], 4 if len(df_raw.columns) > 4 else 0)
    col_fecha = get_col_safe(df_raw, ['fecha', 'date'], 0)

    df_raw[col_camp] = df_raw[col_camp].astype(str).str.strip().str.title()
    df_raw[col_fecha] = pd.to_datetime(df_raw[col_fecha], errors='coerce')
    df_raw = df_raw.dropna(subset=[col_fecha])

    df_raw[col_aht] = df_raw[col_aht].apply(parse_aht_to_seconds)
    df_raw['Total_Segundos_Handle'] = df_raw[col_calls] * df_raw[col_aht]
    df_raw['Inter_Clean'] = df_raw[col_inter].astype(str).str.strip().apply(lambda x: ':'.join(x.split(':'))[:5] if len(x.split(':')) >= 2 else x)

    df = df_raw.groupby([col_fecha, col_camp, 'Inter_Clean']).agg({
        col_calls: 'sum',
        'Total_Segundos_Handle': 'sum'
    }).reset_index()

    df[col_aht] = df.apply(lambda r: r['Total_Segundos_Handle'] / r[col_calls] if r[col_calls] > 0 else 180.0, axis=1)
    df = df.drop(columns=['Total_Segundos_Handle'])

    dias_espanol = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo']
    meses_espanol = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

    df['Dia_Semana_Clean'] = df[col_fecha].dt.weekday.apply(lambda w: dias_espanol[w])

    fecha_maxima = df[col_fecha].max()
    fecha_inicio_forecast = fecha_maxima + timedelta(days=1)
    aht_global_campana = df.groupby(col_camp)[col_aht].apply(lambda x: x[x > 0].mean() if len(x[x > 0]) > 0 else 180.0).to_dict()

    df_diario = df.groupby([col_fecha, col_camp])[col_calls].sum().reset_index()
    campanas_unicas = df[col_camp].unique()

    hw_forecasts = {}
    for camp in campanas_unicas:
        sub = df_diario[df_diario[col_camp] == camp].sort_values(col_fecha).reset_index(drop=True)
        volumenes_list = sub[col_calls].tolist()
        hw_forecasts[camp] = holt_winters_fit_predict_rapido(volumenes_list, n_preds=dias_futuros)

    perfil_intradia = df.groupby([col_camp, 'Dia_Semana_Clean', 'Inter_Clean']).agg(
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
        inters_camp = df[df[col_camp] == camp]['Inter_Clean'].unique().tolist()
        intervalos_operativos_por_camp[camp] = sorted(inters_camp)

    factor_asistencia = max(0.01, 1.0 - merma)
    data_processed = []

    for d in range(dias_futuros):
        fecha_actual = fecha_inicio_forecast + timedelta(days=d)
        str_fecha = fecha_actual.strftime('%Y-%m-%d')
        str_mes = f"{meses_espanol[fecha_actual.month]} {fecha_actual.year}"
        nombre_dia = dias_espanol[fecha_actual.weekday()]

        for camp in campanas_unicas:
            volumen_predicho_diario = hw_forecasts[camp][d]
            intervalos_validos = intervalos_operativos_por_camp.get(camp, [])

            for inter in intervalos_validos:
                key_p = (camp, nombre_dia, inter)
                info_p = mapa_perfil.get(key_p, {'weight': 0.0, 'aht': 0.0})
                calls = volumen_predicho_diario * info_p['weight']
                aht = info_p['aht'] if info_p['aht'] > 0 else aht_global_campana.get(camp, 180.0)

                a_erlang = (calls * aht) / 1800.0 if (aht > 0 and calls > 0) else 0.0
                req_ftes = calcular_agentes_requeridos_erlang_c_rapido(a_erlang, target_sl) if calls > 0 else 0
                req_hc = math.ceil(req_ftes / factor_asistencia) if req_ftes > 0 else 0

                data_processed.append({
                    'Campaña': str(camp),
                    'Fecha': str_fecha,
                    'Mes': str_mes,
                    'Día_Semana': nombre_dia.capitalize(),
                    'Intervalo': inter,
                    'Llamadas': int(round(calls)),
                    'AHT': format_aht_str(aht),
                    'AHT_Segundos': int(round(aht)),
                    'Agentes_Requeridos': req_hc,
                    'HC_Plantilla_Nominal': 29,
                    'HC_Plantilla_Activa_Dia': 24
                })

    DATA_CACHE_IN_MEMORY = data_processed
    gc.collect()
    return data_processed

def resolver_turnos_optimos(intervalos, campanas_activas, llamadas_vec=None, aht_vec=None, 
                            target_sl=80.0, target_time=20.0, merma=0.20, duracion_jornada=6.5, 
                            es_nocturno=True):
    m = len(intervalos)
    if m == 0:
        return [], [0]*m, 0, 0, 100.0, [100.0]*m, 100.0, 100.0, [0]*m

    llamadas_arr = np.nan_to_num(np.array(llamadas_vec, dtype=float), nan=0.0) if llamadas_vec is not None else np.zeros(m)
    aht_arr = np.nan_to_num(np.array(aht_vec, dtype=float), nan=180.0) if aht_vec is not None else np.full(m, 180.0)
    
    tot_llamadas = float(np.sum(llamadas_arr))
    factor_asistencia = max(0.01, 1.0 - merma)
    target_sl_dinamico = float(target_sl)
    
    req_hc_pooled = []
    req_hc_base = np.zeros(m)
    for i in range(m):
        c = llamadas_arr[i]
        aht_s = aht_arr[i]
        a_erl = (c * aht_s) / 1800.0 if (c > 0 and aht_s > 0) else 0.0
        req_ftes_i = calcular_agentes_requeridos_erlang_c_rapido(a_erl, target_sl_dinamico) if c > 0 else 0
        req_hc_i = math.ceil(req_ftes_i / factor_asistencia) if req_ftes_i > 0 else 0
        req_hc_pooled.append(int(req_hc_i))
        req_hc_base[i] = req_hc_i

    cob_hc = np.zeros(m, dtype=float)
    x_turnos_dict = {}

    agentes_nocturnos_totales_hc = 0
    agentes_diurnos_totales_hc = 0

    if es_nocturno:
        label_jornada_noc = "9.0 hrs (Nocturno 5x2)"
        indices_nocturnos = []
        for j in range(m):
            min_in = parse_time_str(intervalos[j])
            if min_in is not None and (min_in >= (22 * 60) or min_in < (7 * 60)):
                indices_nocturnos.append(j)

        if len(indices_nocturnos) > 0:
            agentes_noc_hc = 6
            key_turno_noc = ("22:00", "07:00", label_jornada_noc)
            x_turnos_dict[key_turno_noc] = agentes_noc_hc
            agentes_nocturnos_totales_hc = agentes_noc_hc
            for idx in indices_nocturnos:
                cob_hc[idx] += agentes_noc_hc

    duracion_jornada = float(duracion_jornada)
    SHIFT_BLOCKS = int(round(duracion_jornada * 2))
    duracion_minutos = int(round(duracion_jornada * 60))
    label_jornada_diurna = f"{duracion_jornada:.1f} hrs".replace('.0', '')

    valid_starts = []
    for j in range(m):
        min_in = parse_time_str(intervalos[j])
        if min_in is not None and (7 * 60) <= min_in <= (22 * 60 - duracion_minutos):
            valid_starts.append(j)

    if len(valid_starts) > 0:
        for _ in range(150):
            deficit = req_hc_base - cob_hc
            if np.max(deficit) <= 0: break
            
            best_start_idx = -1
            best_score = -999999
            for s_idx in valid_starts:
                e_idx = min(s_idx + SHIFT_BLOCKS, m)
                sub_deficit = deficit[s_idx:e_idx]
                score = np.sum(np.maximum(0, sub_deficit)) - np.sum(np.maximum(0, -sub_deficit)) * 0.5
                if score > best_score:
                    best_score = score
                    best_start_idx = s_idx

            if best_start_idx == -1 or best_score <= 0: break
                
            min_in_val = parse_time_str(intervalos[best_start_idx])
            min_out_val = min_in_val + duracion_minutos
            h_in_str = f"{(int(min_in_val // 60)):02d}:{(int(min_in_val % 60)):02d}"
            h_out_str = f"{(int(min_out_val // 60)):02d}:{(int(min_out_val % 60)):02d}"
            
            key_turno = (h_in_str, h_out_str, label_jornada_diurna)
            x_turnos_dict[key_turno] = x_turnos_dict.get(key_turno, 0) + 1
            
            for t in range(best_start_idx, min(best_start_idx + SHIFT_BLOCKS, m)):
                cob_hc[t] += 1

    turnos_sugeridos = []
    total_agentes_diarios_hc = 0
    for (h_in, h_out, label_dur), qty in x_turnos_dict.items():
        if qty > 0:
            turnos_sugeridos.append({
                'horario_entrada': h_in,
                'horario_salida': h_out,
                'agentes_a_programar': int(qty),
                'duracion': label_dur
            })
            total_agentes_diarios_hc += int(qty)
            if "Nocturno" not in label_dur:
                agentes_diurnos_totales_hc += int(qty)

    turnos_sugeridos = sorted(turnos_sugeridos, key=lambda x: parse_time_str(x['horario_entrada']) or 0)
    hc_nocturno = math.ceil(agentes_nocturnos_totales_hc * (7.0 / 5.0))
    hc_diurno = math.ceil(agentes_diurnos_totales_hc * (7.0 / 6.0))
    headcount_semanal_requerido = int(hc_nocturno + hc_diurno)

    return turnos_sugeridos, [int(x) for x in cob_hc], total_agentes_diarios_hc, headcount_semanal_requerido, 95.0, [100.0]*m, 98.0, 100.0, [int(x) for x in req_hc_base]

@app.route('/api/optimize-schedules', methods=['POST'])
def api_optimize_schedules():
    try:
        body = request.get_json(force=True)
        intervalos = body.get('intervalos', [])
        campanas = body.get('campanas', [])
        llamadas = [clean_num(x) for x in body.get('llamadas', [])]
        ahts = [clean_num(x, 180.0) for x in body.get('ahts', [])]
        target_sl = float(body.get('target_sl', 80.0))
        target_time = float(body.get('target_time', 20.0))
        merma = float(body.get('merma', 30.0)) / 100.0
        duracion_jornada = float(body.get('duracion_jornada', 6.5))
        es_nocturno = bool(body.get('es_nocturno', True))

        turnos, cob_optima, total_diario, total_hc, eficiencia, sl_vec, sl_global, staff_level, req_hc_pooled = resolver_turnos_optimos(
            intervalos, campanas, llamadas_vec=llamadas, aht_vec=ahts, 
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
            'staffing_level_optimo': staff_level,
            'req_hc_pooled': req_hc_pooled
        }), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/latest', methods=['GET'])
def get_latest_forecast():
    if os.path.exists(EXCEL_DEFAULT):
        try:
            data = procesar_archivo_excel_rapido(EXCEL_DEFAULT, dias_futuros=180)
            return jsonify(data), 200
        except Exception as e:
            return jsonify({'error': f'Error procesando historico.xlsx: {str(e)}'}), 500
    return jsonify({'error': 'No se encontró historico.xlsx'}), 404

@app.route('/api/process', methods=['POST', 'GET'])
def process_data():
    if request.method == 'GET':
        return jsonify({'status': 'API predictiva activa'}), 200

    target_sl = clean_num(request.form.get('target_sl'), 80.0)
    target_time = clean_num(request.form.get('target_time'), 20.0)
    merma = clean_num(request.form.get('merma'), 30.0) / 100.0
    dias_futuros = int(clean_num(request.form.get('dias'), 180))

    file_source = request.files['file'] if 'file' in request.files and request.files['file'].filename != '' else EXCEL_DEFAULT

    try:
        data_processed = procesar_archivo_excel_rapido(file_source, target_sl, target_time, merma, dias_futuros)
        return jsonify(data_processed)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
