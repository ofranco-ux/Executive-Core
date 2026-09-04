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
    return jsonify({"error": "ALERTA CRITICA: No se encontro el archivo index.html."}), 404

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
        if erlang_c_sl_optimizado(A, n, aht, target_time) >= target_sl: return n
        n += 1
    return n

def parse_time_str(t_str):
    if not t_str: return None
    t = re.sub(r'[^\d:]', '', str(t_str).lower())
    if not t: return None
    if ':' not in t: t += ':00'
    try:
        p = t.split(':')
        return int(p[0]) * 60 + int(p[1])
    except: return None

def esta_en_ventana_servicio(campana, intervalo_str):
    camp_key = str(campana).strip().lower()
    min_in = parse_time_str(intervalo_str)
    if min_in is None: return True
    for key, ventana in VENTANAS_SERVICIO.items():
        if key in camp_key or camp_key in key:
            return ventana['inicio'] <= min_in < ventana['fin']
    return True

def encontrar_columna(df, posibles):
    for p in posibles:
        for c in df.columns:
            if p.strip().lower() in str(c).strip().lower(): return c
    return None

def generar_intervalos_cobertura(start_min, end_min):
    intervals = []
    curr = start_min
    if start_min < end_min:
        while curr < end_min:
            intervals.append(f"{int(curr // 60):02d}:{int(curr % 60):02d}")
            curr += 30
    else: 
        while curr < 24 * 60:
            intervals.append(f"{int(curr // 60):02d}:{int(curr % 60):02d}")
            curr += 30
        curr = 0
        while curr < end_min:
            intervals.append(f"{int(curr // 60):02d}:{int(curr % 60):02d}")
            curr += 30
    return intervals

def procesar_hoja_roster(df_roster):
    dias_map = {'lunes': 'Lunes', 'martes': 'Martes', 'miércoles': 'Miércoles', 'miercoles': 'Miércoles', 
                'jueves': 'Jueves', 'viernes': 'Viernes', 'sábado': 'Sábado', 'sabado': 'Sábado', 'domingo': 'Domingo'}
    roster_cov, roster_total_camp, roster_total_dia_camp = {}, {}, {}
    col_camp = encontrar_columna(df_roster, ['campaña', 'campana', 'skill', 'servicio'])
    col_agente = encontrar_columna(df_roster, ['agente', 'nombre', 'asesor', 'ejecutivo', 'id'])
    
    if not col_camp: return roster_cov, roster_total_camp, roster_total_dia_camp
        
    for idx, row in df_roster.iterrows():
        if col_agente:
            agente_val = str(row[col_agente]).strip()
            if agente_val.lower() == 'nan' or agente_val == '': continue
            
        camp = str(row[col_camp]).strip().title()
        if camp == 'Nan' or camp == '': continue
        
        roster_total_camp[camp] = roster_total_camp.get(camp, 0) + 1
        for col in df_roster.columns:
            c_lower = str(col).lower().strip()
            if c_lower in dias_map:
                dia_real = dias_map[c_lower]
                horario = str(row[col]).strip().upper()
                if horario != 'DD-DD' and 'NAN' not in horario and '-' in horario:
                    key_dia = (camp, dia_real)
                    roster_total_dia_camp[key_dia] = roster_total_dia_camp.get(key_dia, 0) + 1
                    parts = horario.split('-')
                    if len(parts) == 2:
                        s_min = parse_time_str(parts[0].strip())
                        e_min = parse_time_str(parts[1].strip())
                        if s_min is not None and e_min is not None:
                            for inv in generar_intervalos_cobertura(s_min, e_min):
                                roster_cov[(camp, dia_real, inv)] = roster_cov.get((camp, dia_real, inv), 0) + 1
    return roster_cov, roster_total_camp, roster_total_dia_camp

def procesar_archivo_excel(file_source, target_sl=80.0, target_time=20.0, merma=0.20, dias_futuros=45):
    xls_file = pd.ExcelFile(file_source, engine='openpyxl')
    sheet_calls = xls_file.sheet_names[0]
    for s in xls_file.sheet_names:
        if 'llam' in s.lower() or 'hist' in s.lower() or 'datos' in s.lower(): sheet_calls = s; break
            
    sheet_roster = None
    for s in xls_file.sheet_names:
        s_lower = s.lower()
        if ('roster' in s_lower or 'plantilla' in s_lower or 'platilla' in s_lower or 'horario' in s_lower) and 'out' not in s_lower and 'salida' not in s_lower and 'chat' not in s_lower and 'mensaje' not in s_lower: 
            sheet_roster = s
            break
            
    if not sheet_roster:
        for s in xls_file.sheet_names:
            if 'roster' in s.lower() or 'plantilla' in s.lower() or 'platilla' in s_lower: 
                sheet_roster = s; break

    roster_coverage, roster_total_camp, roster_total_dia_camp = {}, {}, {}
    if sheet_roster:
        try:
            df_roster = pd.read_excel(xls_file, sheet_name=sheet_roster, engine='openpyxl')
            roster_coverage, roster_total_camp, roster_total_dia_camp = procesar_hoja_roster(df_roster)
        except: pass

    df_raw = pd.read_excel(xls_file, sheet_name=sheet_calls, engine='openpyxl')
    col_calls = encontrar_columna(df_raw, ['recibidas', 'llamadas', 'calls', 'volumen', 'ofrecidas', 'entrada'])
    col_aht = encontrar_columna(df_raw, ['aht', 'tmo', 'handle', 'duracion'])
    col_camp = encontrar_columna(df_raw, ['campaña', 'campana', 'skill', 'servicio', 'ring group'])
    col_inter = encontrar_columna(df_raw, ['intervalo', 'hora', 'time'])
    col_fecha = encontrar_columna(df_raw, ['fecha', 'date'])

    if not col_camp: col_camp = df_raw.columns[0]
    if not col_fecha: col_fecha = df_raw.columns[1]
    if not col_inter: col_inter = df_raw.columns[2]
    if not col_calls: col_calls = df_raw.columns[3]

    df_raw[col_camp] = df_raw[col_camp].astype(str).str.strip().str.title()
    df_raw[col_fecha] = pd.to_datetime(df_raw[col_fecha], dayfirst=True, errors='coerce').dt.normalize()
    df_raw = df_raw.dropna(subset=[col_fecha])
    df_raw[col_calls] = [clean_num(x, 0.0) for x in df_raw[col_calls]]

    df_valido = df_raw[df_raw[col_calls] > 0]
    if df_valido.empty: raise ValueError("El archivo no tiene volumen mayor a cero.")
    
    max_fecha_real = df_valido[col_fecha].max()
    df_raw = df_raw[df_raw[col_fecha] <= max_fecha_real]

    if col_aht: df_raw[col_aht] = [parse_aht_to_seconds(x) for x in df_raw[col_aht]]
    else: df_raw['AHT_Calc'] = 180.0; col_aht = 'AHT_Calc'

    df_raw['Inter_Clean'] = df_raw[col_inter].apply(clean_interval_str)
    df_raw['Total_Segundos_Handle'] = df_raw[col_calls] * df_raw[col_aht]

    df = df_raw.groupby([col_fecha, col_camp, 'Inter_Clean']).agg({col_calls: 'sum', 'Total_Segundos_Handle': 'sum'}).reset_index()
    df[col_aht] = np.where(df[col_calls] > 0, df['Total_Segundos_Handle'] / df[col_calls], 180.0)
    df = df.drop(columns=['Total_Segundos_Handle'])

    dias_espanol = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo']
    meses_espanol = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    df['Dia_Semana_Clean'] = df[col_fecha].dt.weekday.apply(lambda w: dias_espanol[w])

    fecha_inicio_forecast = max_fecha_real + timedelta(days=1)
    aht_global_campana = df.groupby(col_camp)[col_aht].apply(lambda x: x[x > 0].mean() if len(x[x > 0]) > 0 else 180.0).to_dict()

    df_diario = df.groupby([col_fecha, col_camp])[col_calls].sum().reset_index()
    campanas_unicas = df[col_camp].unique()

    predicciones_futuras, factores_ui = {}, {}

    for camp in campanas_unicas:
        sub = df_diario[df_diario[col_camp] == camp].sort_values(col_fecha).reset_index(drop=True)
        if sub.empty: continue
        
        # 1. CLASIFICACIÓN ADAPTATIVA POR VOLUMEN
        vol_promedio = sub[col_calls].mean()
        es_micro_campana = vol_promedio < 200

        # 2. PROMEDIO BASE DINÁMICO
        dow_avg = {}
        for i in range(7):
            vols_dow = sub[sub[col_fecha].dt.weekday == i][col_calls]
            vols_dow = vols_dow[vols_dow > 0]
            vols_list = vols_dow.tail(4).tolist()
            
            if len(vols_list) == 4:
                if es_micro_campana:
                    dow_avg[i] = sum(vols_list) / 4.0
                else:
                    dow_avg[i] = (vols_list[3] * 0.50) + (vols_list[2] * 0.30) + (vols_list[1] * 0.15) + (vols_list[0] * 0.05)
            elif len(vols_list) > 0:
                dow_avg[i] = sum(vols_list) / len(vols_list)
            else:
                dow_avg[i] = vol_promedio

        # 3. FACTOR DE TENDENCIA INTELIGENTE
        fecha_max = sub[col_fecha].max()
        if es_micro_campana:
            ultimos = sub[(sub[col_fecha] > fecha_max - timedelta(days=14))][col_calls].mean()
            previos = sub[(sub[col_fecha] > fecha_max - timedelta(days=28)) & (sub[col_fecha] <= fecha_max - timedelta(days=14))][col_calls].mean()
            tope_min, tope_max = 0.85, 1.15
        else:
            ultimos = sub[(sub[col_fecha] > fecha_max - timedelta(days=7))][col_calls].mean()
            previos = sub[(sub[col_fecha] > fecha_max - timedelta(days=14)) & (sub[col_fecha] <= fecha_max - timedelta(days=7))][col_calls].mean()
            tope_min, tope_max = 0.70, 1.30
        
        ajuste_reciente = 1.0
        if pd.notna(previos) and previos > 0 and pd.notna(ultimos):
            ajuste_reciente = ultimos / previos
            
        ajuste_reciente = max(tope_min, min(tope_max, ajuste_reciente))

        # 4. APLICACIÓN AL FORECAST
        preds_finales = []
        for d in range(dias_futuros):
            fecha_futura = fecha_inicio_forecast + timedelta(days=d)
            wd = fecha_futura.weekday()
            day_m = fecha_futura.day
            
            vol_base = dow_avg.get(wd, sub[col_calls].mean())
            quincena_factor = 1.10 if day_m in [1, 15, 16, 30, 31] else 1.0
            
            vol_final = vol_base * ajuste_reciente * quincena_factor
            if sub[col_calls].mean() < 1.0: vol_final = 0.0
            preds_finales.append(max(0.0, vol_final))

        predicciones_futuras[camp] = preds_finales
        factores_ui[camp] = round(ajuste_reciente, 2)

    df['En_Ventana'] = [esta_en_ventana_servicio(c, i) for c, i in zip(df[col_camp], df['Inter_Clean'])]
    df_filtrado = df[df['En_Ventana']].copy()

    df_reciente = df_filtrado[df_filtrado[col_fecha] >= (max_fecha_real - timedelta(days=28))]
    if df_reciente.empty: df_reciente = df_filtrado.copy()
    
    # PERFIL HÍBRIDO: Día de la semana específico + Respaldo Global si hay huecos
    perfil_intradia = df_reciente.groupby([col_camp, 'Dia_Semana_Clean', 'Inter_Clean']).agg(
        avg_calls=(col_calls, 'mean'),
        avg_aht=(col_aht, lambda x: x[x > 0].mean() if len(x[x > 0]) > 0 else 0)
    ).reset_index()

    totales_dia = perfil_intradia.groupby([col_camp, 'Dia_Semana_Clean'])['avg_calls'].transform('sum')
    perfil_intradia['weight'] = [(c / t) if t > 0 else 0 for c, t in zip(perfil_intradia['avg_calls'], totales_dia)]
    mapa_perfil = {(r[col_camp], r['Dia_Semana_Clean'], r['Inter_Clean']): {'weight': r['weight'], 'aht': r['avg_aht']} for _, r in perfil_intradia.iterrows()}

    # Perfil global de respaldo
    perfil_global = df_reciente.groupby([col_camp, 'Inter_Clean']).agg(avg_calls=(col_calls, 'mean')).reset_index()
    totales_global = perfil_global.groupby([col_camp])['avg_calls'].transform('sum')
    perfil_global['weight'] = [(c / t) if t > 0 else 0 for c, t in zip(perfil_global['avg_calls'], totales_global)]
    mapa_perfil_global = {(r[col_camp], r['Inter_Clean']): r['weight'] for _, r in perfil_global.iterrows()}
    
    todos_los_intervalos_crudos = [f"{int(h):02d}:{int(m):02d}" for h in range(24) for m in (0, 30)]
    intervalos_operativos_por_camp = {camp: [i for i in todos_los_intervalos_crudos if esta_en_ventana_servicio(camp, i)] for camp in campanas_unicas}

    del df_raw, df, df_diario, df_filtrado, df_reciente
    gc.collect()

    factor_asistencia = max(0.01, 1.0 - merma)
    data_processed = []

    for d in range(dias_futuros):
        fecha_actual = fecha_inicio_forecast + timedelta(days=d)
        str_fecha = fecha_actual.strftime('%Y-%m-%d')
        str_mes = f"{meses_espanol[fecha_actual.month]} {fecha_actual.year}"
        nombre_dia = dias_espanol[fecha_actual.weekday()]

        for camp in campanas_unicas:
            vol_diario = predicciones_futuras.get(camp, [0]*dias_futuros)[d]
            factor_visual_ui = factores_ui.get(camp, 1.0)
            intervalos_validos = intervalos_operativos_por_camp.get(camp, [])

            pesos_crudos = []
            for inter in intervalos_validos:
                w = mapa_perfil.get((camp, nombre_dia, inter), {}).get('weight', 0.0)
                if w == 0.0:
                    w = mapa_perfil_global.get((camp, inter), 0.0)
                pesos_crudos.append(w)

            suma_pesos = sum(pesos_crudos)
            if suma_pesos > 0: pesos_norm = [p / suma_pesos for p in pesos_crudos]
            elif len(intervalos_validos) > 0: pesos_norm = [1.0 / len(intervalos_validos)] * len(intervalos_validos)
            else: pesos_norm = []

            exact_calls = [vol_diario * p for p in pesos_norm]
            floor_calls = [int(math.floor(c)) for c in exact_calls]
            remainders = [(exact_calls[i] - floor_calls[i], i) for i in range(len(exact_calls))]
            remainders.sort(reverse=True, key=lambda x: x[0])
            
            diff = int(round(vol_diario)) - sum(floor_calls)
            for i in range(diff):
                if i < len(remainders): floor_calls[remainders[i][1]] += 1

            aht_global = aht_global_campana.get(camp, 180.0)
            for idx_inter, inter in enumerate(intervalos_validos):
                calls_int = floor_calls[idx_inter]
                calls_float = exact_calls[idx_inter] 

                info_p = mapa_perfil.get((camp, nombre_dia, inter), {})
                aht_real = info_p.get('aht', 0.0)
                
                if aht_real > 0 and not pd.isna(aht_real):
                    aht = aht_real if aht_real >= (aht_global * 0.5) else aht_global
                else:
                    aht = aht_global

                if calls_int <= 0:
                    aht = 0.0

                a_erlang = (calls_float * aht) / 1800.0 if (aht > 0 and calls_float > 0) else 0.0
                req_ftes = calcular_agentes_requeridos_erlang_c(a_erlang, aht, target_time, target_sl) if calls_float > 0 else 0
                req_hc = math.ceil(req_ftes / factor_asistencia) if req_ftes > 0 else 0
                
                hc_roster = roster_coverage.get((str(camp), nombre_dia.capitalize(), inter), 0)
                tot_camp = roster_total_camp.get(str(camp), 0)
                tot_camp_dia = roster_total_dia_camp.get((str(camp), nombre_dia.capitalize()), 0)

                data_processed.append({
                    'Campaña': str(camp),
                    'Fecha': str_fecha,
                    'Mes': str_mes,
                    'Día_Semana': nombre_dia.capitalize(),
                    'Intervalo': inter,
                    'Llamadas': calls_int,
                    'AHT': format_aht_str(aht),
                    'AHT_Segundos': int(round(aht)),
                    'Agentes_Requeridos': req_hc,
                    'HC_Actual_Roster': hc_roster,
                    'Total_Roster_Campana': tot_camp,
                    'Total_Roster_Dia': tot_camp_dia,
                    'Factor_Correccion': factor_visual_ui
                })

    try:
        with open(CACHE_FILE_IN, 'w', encoding='utf-8') as f: json.dump(data_processed, f)
    except: pass
    return data_processed

def procesar_archivo_outbound(file_source, merma=0.20, dias_futuros=45):
    xls_file = pd.ExcelFile(file_source, engine='openpyxl')
    
    sheet_out = None
    for s in xls_file.sheet_names:
        if 'out' in s.lower() or 'salida' in s.lower(): 
            sheet_out = s
            break
            
    if not sheet_out:
        raise ValueError("No se encontró una pestaña llamada 'Out' o 'Salida' en el archivo Excel para procesar Outbound.")

    sheet_roster = None
    for s in xls_file.sheet_names:
        s_lower = s.lower()
        if ('plantilla' in s_lower or 'platilla' in s_lower or 'roster' in s_lower) and ('out' in s_lower or 'salida' in s_lower):
            sheet_roster = s
            break

    roster_coverage, roster_total_camp, roster_total_dia_camp = {}, {}, {}
    if sheet_roster:
        try:
            df_roster = pd.read_excel(xls_file, sheet_name=sheet_roster, engine='openpyxl')
            roster_coverage, roster_total_camp, roster_total_dia_camp = procesar_hoja_roster(df_roster)
        except: pass

    df_raw = pd.read_excel(xls_file, sheet_name=sheet_out, engine='openpyxl')
    col_calls = encontrar_columna(df_raw, ['realizadas', 'llamadas', 'out'])
    col_aht = encontrar_columna(df_raw, ['aht', 'tmo', 'handle', 'duracion'])
    col_camp = encontrar_columna(df_raw, ['campaña', 'campana', 'skill'])
    col_inter = encontrar_columna(df_raw, ['intervalo', 'hora', 'time'])
    col_fecha = encontrar_columna(df_raw, ['fecha', 'date'])

    if not col_camp: col_camp = df_raw.columns[3]
    if not col_fecha: col_fecha = df_raw.columns[0]
    if not col_inter: col_inter = df_raw.columns[4]
    if not col_calls: col_calls = df_raw.columns[5]

    df_raw[col_camp] = df_raw[col_camp].astype(str).str.strip().str.title()
    df_raw[col_fecha] = pd.to_datetime(df_raw[col_fecha], dayfirst=True, errors='coerce').dt.normalize()
    df_raw = df_raw.dropna(subset=[col_fecha])
    df_raw[col_calls] = [clean_num(x, 0.0) for x in df_raw[col_calls]]

    df_valido = df_raw[df_raw[col_calls] > 0]
    if df_valido.empty: raise ValueError("El archivo Outbound no tiene volumen mayor a cero.")
    
    max_fecha_real = df_valido[col_fecha].max()
    df_raw = df_raw[df_raw[col_fecha] <= max_fecha_real]

    if col_aht: df_raw[col_aht] = [parse_aht_to_seconds(x) for x in df_raw[col_aht]]
    else: df_raw['AHT_Calc'] = 180.0; col_aht = 'AHT_Calc'

    df_raw['Inter_Clean'] = df_raw[col_inter].apply(clean_interval_str)
    df_raw['Total_Segundos_Handle'] = df_raw[col_calls] * df_raw[col_aht]

    df = df_raw.groupby([col_fecha, col_camp, 'Inter_Clean']).agg({col_calls: 'sum', 'Total_Segundos_Handle': 'sum'}).reset_index()
    df[col_aht] = np.where(df[col_calls] > 0, df['Total_Segundos_Handle'] / df[col_calls], 180.0)
    df = df.drop(columns=['Total_Segundos_Handle'])

    dias_espanol = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo']
    meses_espanol = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    df['Dia_Semana_Clean'] = df[col_fecha].dt.weekday.apply(lambda w: dias_espanol[w])

    fecha_inicio_forecast = max_fecha_real + timedelta(days=1)
    aht_global_campana = df.groupby(col_camp)[col_aht].apply(lambda x: x[x > 0].mean() if len(x[x > 0]) > 0 else 180.0).to_dict()

    df_diario = df.groupby([col_fecha, col_camp])[col_calls].sum().reset_index()
    campanas_unicas = df[col_camp].unique()

    predicciones_futuras, factores_ui = {}, {}

    for camp in campanas_unicas:
        sub = df_diario[df_diario[col_camp] == camp].sort_values(col_fecha).reset_index(drop=True)
        if sub.empty: continue
        
        # 1. CLASIFICACIÓN ADAPTATIVA POR VOLUMEN
        vol_promedio = sub[col_calls].mean()
        es_micro_campana = vol_promedio < 200

        # 2. PROMEDIO BASE DINÁMICO
        dow_avg = {}
        for i in range(7):
            vols_dow = sub[sub[col_fecha].dt.weekday == i][col_calls]
            vols_dow = vols_dow[vols_dow > 0]
            vols_list = vols_dow.tail(4).tolist()
            
            if len(vols_list) == 4:
                if es_micro_campana:
                    dow_avg[i] = sum(vols_list) / 4.0
                else:
                    dow_avg[i] = (vols_list[3] * 0.50) + (vols_list[2] * 0.30) + (vols_list[1] * 0.15) + (vols_list[0] * 0.05)
            elif len(vols_list) > 0:
                dow_avg[i] = sum(vols_list) / len(vols_list)
            else:
                dow_avg[i] = vol_promedio

        # 3. FACTOR DE TENDENCIA INTELIGENTE
        fecha_max = sub[col_fecha].max()
        if es_micro_campana:
            ultimos = sub[(sub[col_fecha] > fecha_max - timedelta(days=14))][col_calls].mean()
            previos = sub[(sub[col_fecha] > fecha_max - timedelta(days=28)) & (sub[col_fecha] <= fecha_max - timedelta(days=14))][col_calls].mean()
            tope_min, tope_max = 0.85, 1.15
        else:
            ultimos = sub[(sub[col_fecha] > fecha_max - timedelta(days=7))][col_calls].mean()
            previos = sub[(sub[col_fecha] > fecha_max - timedelta(days=14)) & (sub[col_fecha] <= fecha_max - timedelta(days=7))][col_calls].mean()
            tope_min, tope_max = 0.70, 1.30
        
        ajuste_reciente = 1.0
        if pd.notna(previos) and previos > 0 and pd.notna(ultimos):
            ajuste_reciente = ultimos / previos
            
        ajuste_reciente = max(tope_min, min(tope_max, ajuste_reciente))

        # 4. APLICACIÓN AL FORECAST
        preds_finales = []
        for d in range(dias_futuros):
            fecha_futura = fecha_inicio_forecast + timedelta(days=d)
            wd = fecha_futura.weekday()
            day_m = fecha_futura.day
            
            vol_base = dow_avg.get(wd, sub[col_calls].mean())
            quincena_factor = 1.10 if day_m in [1, 15, 16, 30, 31] else 1.0
            
            vol_final = vol_base * ajuste_reciente * quincena_factor
            if sub[col_calls].mean() < 1.0: vol_final = 0.0
            preds_finales.append(max(0.0, vol_final))

        predicciones_futuras[camp] = preds_finales
        factores_ui[camp] = round(ajuste_reciente, 2)

    df['En_Ventana'] = [esta_en_ventana_servicio(c, i) for c, i in zip(df[col_camp], df['Inter_Clean'])]
    df_filtrado = df[df['En_Ventana']].copy()

    df_reciente = df_filtrado[df_filtrado[col_fecha] >= (max_fecha_real - timedelta(days=28))]
    if df_reciente.empty: df_reciente = df_filtrado.copy()
    
    # PERFIL HÍBRIDO: Día de la semana específico + Respaldo Global si hay huecos
    perfil_intradia = df_reciente.groupby([col_camp, 'Dia_Semana_Clean', 'Inter_Clean']).agg(
        avg_calls=(col_calls, 'mean'),
        avg_aht=(col_aht, lambda x: x[x > 0].mean() if len(x[x > 0]) > 0 else 0)
    ).reset_index()

    totales_dia = perfil_intradia.groupby([col_camp, 'Dia_Semana_Clean'])['avg_calls'].transform('sum')
    perfil_intradia['weight'] = [(c / t) if t > 0 else 0 for c, t in zip(perfil_intradia['avg_calls'], totales_dia)]
    mapa_perfil = {(r[col_camp], r['Dia_Semana_Clean'], r['Inter_Clean']): {'weight': r['weight'], 'aht': r['avg_aht']} for _, r in perfil_intradia.iterrows()}

    # Perfil global de respaldo
    perfil_global = df_reciente.groupby([col_camp, 'Inter_Clean']).agg(avg_calls=(col_calls, 'mean')).reset_index()
    totales_global = perfil_global.groupby([col_camp])['avg_calls'].transform('sum')
    perfil_global['weight'] = [(c / t) if t > 0 else 0 for c, t in zip(perfil_global['avg_calls'], totales_global)]
    mapa_perfil_global = {(r[col_camp], r['Inter_Clean']): r['weight'] for _, r in perfil_global.iterrows()}
    
    todos_los_intervalos_crudos = [f"{int(h):02d}:{int(m):02d}" for h in range(24) for m in (0, 30)]
    intervalos_operativos_por_camp = {camp: [i for i in todos_los_intervalos_crudos if esta_en_ventana_servicio(camp, i)] for camp in campanas_unicas}

    del df_raw, df, df_diario, df_filtrado, df_reciente
    gc.collect()

    factor_asistencia = max(0.01, 1.0 - merma)
    data_processed = []

    for d in range(dias_futuros):
        fecha_actual = fecha_inicio_forecast + timedelta(days=d)
        str_fecha = fecha_actual.strftime('%Y-%m-%d')
        str_mes = f"{meses_espanol[fecha_actual.month]} {fecha_actual.year}"
        nombre_dia = dias_espanol[fecha_actual.weekday()]

        for camp in campanas_unicas:
            vol_diario = predicciones_futuras.get(camp, [0]*dias_futuros)[d]
            factor_visual_ui = factores_ui.get(camp, 1.0)
            intervalos_validos = intervalos_operativos_por_camp.get(camp, [])

            pesos_crudos = []
            for inter in intervalos_validos:
                w = mapa_perfil.get((camp, nombre_dia, inter), {}).get('weight', 0.0)
                if w == 0.0:
                    w = mapa_perfil_global.get((camp, inter), 0.0)
                pesos_crudos.append(w)

            suma_pesos = sum(pesos_crudos)
            if suma_pesos > 0: pesos_norm = [p / suma_pesos for p in pesos_crudos]
            elif len(intervalos_validos) > 0: pesos_norm = [1.0 / len(intervalos_validos)] * len(intervalos_validos)
            else: pesos_norm = []

            exact_calls = [vol_diario * p for p in pesos_norm]
            floor_calls = [int(math.floor(c)) for c in exact_calls]
            remainders = [(exact_calls[i] - floor_calls[i], i) for i in range(len(exact_calls))]
            remainders.sort(reverse=True, key=lambda x: x[0])
            
            diff = int(round(vol_diario)) - sum(floor_calls)
            for i in range(diff):
                if i < len(remainders): floor_calls[remainders[i][1]] += 1

            aht_global = aht_global_campana.get(camp, 180.0)
            for idx_inter, inter in enumerate(intervalos_validos):
                calls_int = floor_calls[idx_inter]
                calls_float = exact_calls[idx_inter] 

                info_p = mapa_perfil.get((camp, nombre_dia, inter), {})
                aht_real = info_p.get('aht', 0.0)
                
                if aht_real > 0 and not pd.isna(aht_real):
                    aht = aht_real if aht_real >= (aht_global * 0.5) else aht_global
                else:
                    aht = aht_global

                if calls_int <= 0:
                    aht = 0.0

                req_ftes = (calls_float * aht) / 1800.0 if (aht > 0 and calls_float > 0) else 0.0
                req_hc = math.ceil(req_ftes / factor_asistencia) if req_ftes > 0 else 0
                
                hc_roster = roster_coverage.get((str(camp), nombre_dia.capitalize(), inter), 0)
                tot_camp = roster_total_camp.get(str(camp), 0)
                tot_camp_dia = roster_total_dia_camp.get((str(camp), nombre_dia.capitalize()), 0)

                data_processed.append({
                    'Campaña': str(camp),
                    'Fecha': str_fecha,
                    'Mes': str_mes,
                    'Día_Semana': nombre_dia.capitalize(),
                    'Intervalo': inter,
                    'Llamadas': calls_int,
                    'AHT': format_aht_str(aht),
                    'AHT_Segundos': int(round(aht)),
                    'Agentes_Requeridos': req_hc,
                    'HC_Actual_Roster': hc_roster,
                    'Total_Roster_Campana': tot_camp,
                    'Total_Roster_Dia': tot_camp_dia,
                    'Factor_Correccion': factor_visual_ui
                })

    try:
        with open(CACHE_FILE_OUT, 'w', encoding='utf-8') as f: json.dump(data_processed, f)
    except: pass
    return data_processed

def procesar_archivo_chat(file_source, target_sl=80.0, target_time=20.0, merma=0.20, concurrencia=3.0, dias_futuros=45):
    xls_file = pd.ExcelFile(file_source, engine='openpyxl')
    
    sheet_chat = None
    for s in xls_file.sheet_names:
        if ('chat' in s.lower() or 'mensaje' in s.lower()) and ('plantilla' not in s.lower() and 'roster' not in s.lower() and 'platilla' not in s.lower()): 
            sheet_chat = s
            break
            
    if not sheet_chat:
        raise ValueError("No se encontró una pestaña llamada 'Chat' o 'Mensajes' en el archivo Excel.")

    sheet_roster = None
    for s in xls_file.sheet_names:
        s_lower = s.lower()
        if ('plantilla' in s_lower or 'platilla' in s_lower or 'roster' in s_lower) and ('chat' in s_lower or 'mensaje' in s_lower):
            sheet_roster = s
            break

    roster_coverage, roster_total_camp, roster_total_dia_camp = {}, {}, {}
    if sheet_roster:
        try:
            df_roster = pd.read_excel(xls_file, sheet_name=sheet_roster, engine='openpyxl')
            roster_coverage, roster_total_camp, roster_total_dia_camp = procesar_hoja_roster(df_roster)
        except: pass

    df_raw = pd.read_excel(xls_file, sheet_name=sheet_chat, engine='openpyxl')
    col_calls = encontrar_columna(df_raw, ['recibidos', 'recibidas', 'llamadas', 'chats', 'mensajes'])
    col_aht = encontrar_columna(df_raw, ['aht', 'tmo', 'handle', 'duracion'])
    col_camp = encontrar_columna(df_raw, ['campaña', 'campana', 'skill'])
    col_inter = encontrar_columna(df_raw, ['intervalo', 'hora', 'time'])
    col_fecha = encontrar_columna(df_raw, ['fecha', 'date'])

    if not col_camp: col_camp = df_raw.columns[3]
    if not col_fecha: col_fecha = df_raw.columns[0]
    if not col_inter: col_inter = df_raw.columns[4]
    if not col_calls: col_calls = df_raw.columns[5]

    df_raw[col_camp] = df_raw[col_camp].astype(str).str.strip().str.title()
    df_raw[col_fecha] = pd.to_datetime(df_raw[col_fecha], dayfirst=True, errors='coerce').dt.normalize()
    df_raw = df_raw.dropna(subset=[col_fecha])
    df_raw[col_calls] = [clean_num(x, 0.0) for x in df_raw[col_calls]]

    df_valido = df_raw[df_raw[col_calls] > 0]
    if df_valido.empty: raise ValueError("El archivo Chat no tiene volumen mayor a cero.")
    
    max_fecha_real = df_valido[col_fecha].max()
    df_raw = df_raw[df_raw[col_fecha] <= max_fecha_real]

    if col_aht: df_raw[col_aht] = [parse_aht_to_seconds(x) for x in df_raw[col_aht]]
    else: df_raw['AHT_Calc'] = 600.0; col_aht = 'AHT_Calc'

    df_raw['Inter_Clean'] = df_raw[col_inter].apply(clean_interval_str)
    df_raw['Total_Segundos_Handle'] = df_raw[col_calls] * df_raw[col_aht]

    df = df_raw.groupby([col_fecha, col_camp, 'Inter_Clean']).agg({col_calls: 'sum', 'Total_Segundos_Handle': 'sum'}).reset_index()
    df[col_aht] = np.where(df[col_calls] > 0, df['Total_Segundos_Handle'] / df[col_calls], 600.0)
    df = df.drop(columns=['Total_Segundos_Handle'])

    dias_espanol = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo']
    meses_espanol = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    df['Dia_Semana_Clean'] = df[col_fecha].dt.weekday.apply(lambda w: dias_espanol[w])

    fecha_inicio_forecast = max_fecha_real + timedelta(days=1)
    aht_global_campana = df.groupby(col_camp)[col_aht].apply(lambda x: x[x > 0].mean() if len(x[x > 0]) > 0 else 600.0).to_dict()

    df_diario = df.groupby([col_fecha, col_camp])[col_calls].sum().reset_index()
    campanas_unicas = df[col_camp].unique()

    predicciones_futuras, factores_ui = {}, {}

    for camp in campanas_unicas:
        sub = df_diario[df_diario[col_camp] == camp].sort_values(col_fecha).reset_index(drop=True)
        if sub.empty: continue
        
        # 1. CLASIFICACIÓN ADAPTATIVA POR VOLUMEN
        vol_promedio = sub[col_calls].mean()
        es_micro_campana = vol_promedio < 200

        # 2. PROMEDIO BASE DINÁMICO
        dow_avg = {}
        for i in range(7):
            vols_dow = sub[sub[col_fecha].dt.weekday == i][col_calls]
            vols_dow = vols_dow[vols_dow > 0]
            vols_list = vols_dow.tail(4).tolist()
            
            if len(vols_list) == 4:
                if es_micro_campana:
                    dow_avg[i] = sum(vols_list) / 4.0
                else:
                    dow_avg[i] = (vols_list[3] * 0.50) + (vols_list[2] * 0.30) + (vols_list[1] * 0.15) + (vols_list[0] * 0.05)
            elif len(vols_list) > 0:
                dow_avg[i] = sum(vols_list) / len(vols_list)
            else:
                dow_avg[i] = vol_promedio

        # 3. FACTOR DE TENDENCIA INTELIGENTE
        fecha_max = sub[col_fecha].max()
        if es_micro_campana:
            ultimos = sub[(sub[col_fecha] > fecha_max - timedelta(days=14))][col_calls].mean()
            previos = sub[(sub[col_fecha] > fecha_max - timedelta(days=28)) & (sub[col_fecha] <= fecha_max - timedelta(days=14))][col_calls].mean()
            tope_min, tope_max = 0.85, 1.15
        else:
            ultimos = sub[(sub[col_fecha] > fecha_max - timedelta(days=7))][col_calls].mean()
            previos = sub[(sub[col_fecha] > fecha_max - timedelta(days=14)) & (sub[col_fecha] <= fecha_max - timedelta(days=7))][col_calls].mean()
            tope_min, tope_max = 0.70, 1.30
        
        ajuste_reciente = 1.0
        if pd.notna(previos) and previos > 0 and pd.notna(ultimos):
            ajuste_reciente = ultimos / previos
            
        ajuste_reciente = max(tope_min, min(tope_max, ajuste_reciente))

        # 4. APLICACIÓN AL FORECAST
        preds_finales = []
        for d in range(dias_futuros):
            fecha_futura = fecha_inicio_forecast + timedelta(days=d)
            wd = fecha_futura.weekday()
            day_m = fecha_futura.day
            
            vol_base = dow_avg.get(wd, sub[col_calls].mean())
            quincena_factor = 1.10 if day_m in [1, 15, 16, 30, 31] else 1.0
            
            vol_final = vol_base * ajuste_reciente * quincena_factor
            if sub[col_calls].mean() < 1.0: vol_final = 0.0
            preds_finales.append(max(0.0, vol_final))

        predicciones_futuras[camp] = preds_finales
        factores_ui[camp] = round(ajuste_reciente, 2)

    df['En_Ventana'] = [esta_en_ventana_servicio(c, i) for c, i in zip(df[col_camp], df['Inter_Clean'])]
    df_filtrado = df[df['En_Ventana']].copy()

    df_reciente = df_filtrado[df_filtrado[col_fecha] >= (max_fecha_real - timedelta(days=28))]
    if df_reciente.empty: df_reciente = df_filtrado.copy()
    
    # PERFIL HÍBRIDO: Día de la semana específico + Respaldo Global si hay huecos
    perfil_intradia = df_reciente.groupby([col_camp, 'Dia_Semana_Clean', 'Inter_Clean']).agg(
        avg_calls=(col_calls, 'mean'),
        avg_aht=(col_aht, lambda x: x[x > 0].mean() if len(x[x > 0]) > 0 else 0)
    ).reset_index()

    totales_dia = perfil_intradia.groupby([col_camp, 'Dia_Semana_Clean'])['avg_calls'].transform('sum')
    perfil_intradia['weight'] = [(c / t) if t > 0 else 0 for c, t in zip(perfil_intradia['avg_calls'], totales_dia)]
    mapa_perfil = {(r[col_camp], r['Dia_Semana_Clean'], r['Inter_Clean']): {'weight': r['weight'], 'aht': r['avg_aht']} for _, r in perfil_intradia.iterrows()}

    # Perfil global de respaldo
    perfil_global = df_reciente.groupby([col_camp, 'Inter_Clean']).agg(avg_calls=(col_calls, 'mean')).reset_index()
    totales_global = perfil_global.groupby([col_camp])['avg_calls'].transform('sum')
    perfil_global['weight'] = [(c / t) if t > 0 else 0 for c, t in zip(perfil_global['avg_calls'], totales_global)]
    mapa_perfil_global = {(r[col_camp], r['Inter_Clean']): r['weight'] for _, r in perfil_global.iterrows()}
    
    todos_los_intervalos_crudos = [f"{int(h):02d}:{int(m):02d}" for h in range(24) for m in (0, 30)]
    intervalos_operativos_por_camp = {camp: [i for i in todos_los_intervalos_crudos if esta_en_ventana_servicio(camp, i)] for camp in campanas_unicas}

    del df_raw, df, df_diario, df_filtrado, df_reciente
    gc.collect()

    factor_asistencia = max(0.01, 1.0 - merma)
    data_processed = []

    for d in range(dias_futuros):
        fecha_actual = fecha_inicio_forecast + timedelta(days=d)
        str_fecha = fecha_actual.strftime('%Y-%m-%d')
        str_mes = f"{meses_espanol[fecha_actual.month]} {fecha_actual.year}"
        nombre_dia = dias_espanol[fecha_actual.weekday()]

        for camp in campanas_unicas:
            vol_diario = predicciones_futuras.get(camp, [0]*dias_futuros)[d]
            factor_visual_ui = factores_ui.get(camp, 1.0)
            intervalos_validos = intervalos_operativos_por_camp.get(camp, [])

            pesos_crudos = []
            for inter in intervalos_validos:
                w = mapa_perfil.get((camp, nombre_dia, inter), {}).get('weight', 0.0)
                if w == 0.0:
                    w = mapa_perfil_global.get((camp, inter), 0.0)
                pesos_crudos.append(w)

            suma_pesos = sum(pesos_crudos)
            if suma_pesos > 0: pesos_norm = [p / suma_pesos for p in pesos_crudos]
            elif len(intervalos_validos) > 0: pesos_norm = [1.0 / len(intervalos_validos)] * len(intervalos_validos)
            else: pesos_norm = []

            exact_calls = [vol_diario * p for p in pesos_norm]
            floor_calls = [int(math.floor(c)) for c in exact_calls]
            remainders = [(exact_calls[i] - floor_calls[i], i) for i in range(len(exact_calls))]
            remainders.sort(reverse=True, key=lambda x: x[0])
            
            diff = int(round(vol_diario)) - sum(floor_calls)
            for i in range(diff):
                if i < len(remainders): floor_calls[remainders[i][1]] += 1

            aht_global = aht_global_campana.get(camp, 600.0)
            for idx_inter, inter in enumerate(intervalos_validos):
                calls_int = floor_calls[idx_inter]
                calls_float = exact_calls[idx_inter] 

                info_p = mapa_perfil.get((camp, nombre_dia, inter), {})
                aht_real = info_p.get('aht', 0.0)
                
                if aht_real > 0 and not pd.isna(aht_real):
                    aht = aht_real if aht_real >= (aht_global * 0.5) else aht_global
                else:
                    aht = aht_global

                if calls_int <= 0:
                    aht = 0.0

                aht_efectivo = aht / max(1.0, concurrencia)
                a_erlang_raw = (calls_float * aht_efectivo) / 1800.0 if (aht_efectivo > 0 and calls_float > 0) else 0.0
                
                req_ftes = calcular_agentes_requeridos_erlang_c(a_erlang_raw, aht_efectivo, target_time, target_sl) if calls_float > 0 else 0
                req_hc = math.ceil(req_ftes / factor_asistencia) if req_ftes > 0 else 0.0

                hc_roster = roster_coverage.get((str(camp), nombre_dia.capitalize(), inter), 0)
                tot_camp = roster_total_camp.get(str(camp), 0)
                tot_camp_dia = roster_total_dia_camp.get((str(camp), nombre_dia.capitalize()), 0)

                data_processed.append({
                    'Campaña': str(camp),
                    'Fecha': str_fecha,
                    'Mes': str_mes,
                    'Día_Semana': nombre_dia.capitalize(),
                    'Intervalo': inter,
                    'Llamadas': calls_int,
                    'AHT': format_aht_str(aht),
                    'AHT_Segundos': int(round(aht)),
                    'Agentes_Requeridos': req_hc,
                    'HC_Actual_Roster': hc_roster,
                    'Total_Roster_Campana': tot_camp,
                    'Total_Roster_Dia': tot_camp_dia,
                    'Factor_Correccion': factor_visual_ui
                })

    try:
        with open(CACHE_FILE_CHAT, 'w', encoding='utf-8') as f: json.dump(data_processed, f)
    except: pass
    return data_processed

@app.route('/api/latest', methods=['GET'])
def get_latest_forecast():
    mode = request.args.get('mode', 'inbound')
    if mode == 'outbound': target_cache = CACHE_FILE_OUT
    elif mode == 'chat': target_cache = CACHE_FILE_CHAT
    else: target_cache = CACHE_FILE_IN
    
    if os.path.exists(target_cache):
        try:
            with open(target_cache, 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
                if isinstance(cache_data, list) and len(cache_data) > 0: 
                    return jsonify(cache_data), 200
        except Exception as e:
            print(f"Error leyendo caché: {e}")
            pass
            
    excel_path = buscar_archivo_excel()
    if excel_path:
        try:
            sl, tt, merma, dias, concurrencia = 80.0, 20.0, 30.0, 130, 3.0
            if os.path.exists(CONFIG_FILE):
                try:
                    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                        cfg = json.load(f)
                        sl = float(cfg.get('targetSl', 80.0))
                        tt = float(cfg.get('targetTime', 20.0))
                        
                        # Extraer merma si está en el archivo config
                        merma_data = cfg.get('merma', 30.0)
                        if isinstance(merma_data, dict):
                            merma = float(merma_data.get(mode, 30.0))
                        else:
                            merma = float(merma_data)
                except: pass
            
            merma_pct = merma / 100.0
            
            if mode == 'outbound':
                data = procesar_archivo_outbound(excel_path, merma=merma_pct, dias_futuros=dias)
            elif mode == 'chat':
                data = procesar_archivo_chat(excel_path, target_sl=sl, target_time=tt, merma=merma_pct, concurrencia=concurrencia, dias_futuros=dias)
            else:
                data = procesar_archivo_excel(excel_path, target_sl=sl, target_time=tt, merma=merma_pct, dias_futuros=dias)
                
            gc.collect()
            return jsonify(data), 200
        except Exception as e:
            print(f"Error en auto-recovery ({mode}): {e}")
            pass

    return jsonify([]), 200

@app.route('/api/process', methods=['POST', 'GET'])
def process_data():
    if request.method == 'GET': return jsonify({'status': 'API activa'}), 200
    excel_path = buscar_archivo_excel()
    if not excel_path: return jsonify({'error': 'No se encontro Excel (.xlsx).'}), 400
    try:
        data = procesar_archivo_excel(
            excel_path, 
            float(clean_num(request.form.get('target_sl'), 80.0)), 
            float(clean_num(request.form.get('target_time'), 20.0)), 
            float(clean_num(request.form.get('merma'), 30.0)) / 100.0, 
            int(clean_num(request.form.get('dias'), 45))
        )
        gc.collect()
        return jsonify(data)
    except Exception as e:
        gc.collect()
        return jsonify({'error': str(e)}), 500

@app.route('/api/process_outbound', methods=['POST'])
def process_outbound_data():
    excel_path = buscar_archivo_excel()
    if not excel_path: return jsonify({'error': 'No se encontro Excel (.xlsx).'}), 400
    try:
        data = procesar_archivo_outbound(
            excel_path, 
            float(clean_num(request.form.get('merma'), 30.0)) / 100.0, 
            int(clean_num(request.form.get('dias'), 45))
        )
        gc.collect()
        return jsonify(data)
    except Exception as e:
        gc.collect()
        return jsonify({'error': str(e)}), 500

@app.route('/api/process_chat', methods=['POST'])
def process_chat_data():
    excel_path = buscar_archivo_excel()
    if not excel_path: return jsonify({'error': 'No se encontro Excel (.xlsx).'}), 400
    try:
        data = procesar_archivo_chat(
            excel_path, 
            float(clean_num(request.form.get('target_sl'), 80.0)),
            float(clean_num(request.form.get('target_time'), 20.0)),
            float(clean_num(request.form.get('merma'), 30.0)) / 100.0, 
            float(clean_num(request.form.get('concurrencia'), 3.0)), 
            int(clean_num(request.form.get('dias'), 45))
        )
        gc.collect()
        return jsonify(data)
    except Exception as e:
        gc.collect()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
