from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd
import math
import datetime
from datetime import timedelta
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)
CORS(app)

@app.route('/')
def home():
    return send_from_directory(BASE_DIR, 'index.html')

# ==========================================
# MOTOR MATEMÁTICO (ERLANG C & SL REAL)
# ==========================================
def calcular_probabilidad_espera(trafico_erlangs, agentes):
    if agentes <= trafico_erlangs: return 1.0 
    inv_b = 1.0
    for i in range(1, int(agentes) + 1):
        inv_b = 1.0 + (i / trafico_erlangs) * inv_b
    erlang_b = 1.0 / inv_b
    return (agentes * erlang_b) / (agentes - trafico_erlangs * (1.0 - erlang_b))

def calcular_sl_logrado(volumen, aht, target_time, agentes_reales):
    if volumen <= 0 or aht <= 0 or agentes_reales <= 0: return 0.0
    trafico_erlangs = (volumen * aht) / 1800.0
    if agentes_reales <= trafico_erlangs: return 0.0
    
    prob_espera = calcular_probabilidad_espera(trafico_erlangs, agentes_reales)
    exponente = -(agentes_reales - trafico_erlangs) * (target_time / aht)
    sl = (1.0 - (prob_espera * math.exp(exponente))) * 100.0
    return max(0.0, min(100.0, sl))

def calcular_requerimiento(volumen, aht, target_sl, target_time, merma):
    if volumen <= 0 or aht <= 0: return 0, 0
    trafico_erlangs = (volumen * aht) / 1800.0 
    if trafico_erlangs <= 0: return 0, 0
    
    agentes = math.floor(trafico_erlangs) + 1
    while True:
        prob_espera = calcular_probabilidad_espera(trafico_erlangs, agentes)
        exponente = -(agentes - trafico_erlangs) * (target_time / aht)
        sl_logrado = 1.0 - (prob_espera * math.exp(exponente))
        if sl_logrado >= target_sl: break
        agentes += 1
        
    agentes_programados = math.ceil(agentes / (1.0 - merma))
    return agentes, agentes_programados

def normalizar_intervalo(val):
    if pd.isna(val): return ""
    if isinstance(val, (datetime.time, datetime.datetime)):
        return val.strftime('%H:%M')
    val_str = str(val).strip()
    if len(val_str) == 4 and ':' in val_str:
        val_str = '0' + val_str
    return val_str[:5]

def analizar_dia_historico_campana(df_campana, nombre_dia):
    df_dia = df_campana[df_campana['Día'] == nombre_dia]
    if df_dia.empty: return 0, {}, {} 

    volumen_esperado = int(df_dia.groupby('Fecha')['Recibidas'].sum().mean())
    agrupado = df_dia.groupby('Intervalo').agg(
        Total_Recibidas=('Recibidas', 'sum'),
        Suma_AHT=('Recibidas', lambda x: (x * df_dia.loc[x.index, 'AHT']).sum())
    ).reset_index()
    
    total_dia = agrupado['Total_Recibidas'].sum()
    if total_dia == 0: return 0, {}, {}
    
    curva_vol = dict(zip(agrupado['Intervalo'], agrupado['Total_Recibidas'] / total_dia))
    curva_aht = dict(zip(agrupado['Intervalo'], (agrupado['Suma_AHT'] / agrupado['Total_Recibidas']).fillna(0)))
    return volumen_esperado, curva_vol, curva_aht

def contar_agentes_desde_malla(df_agentes, campana, nombre_dia_columna, intervalo):
    if df_agentes.empty or nombre_dia_columna not in df_agentes.columns:
        return 0
        
    agentes_campana = df_agentes[df_agentes['Campaña'].astype(str).str.strip().str.lower() == str(campana).strip().lower()]
    agentes_activos = 0
    for horario in agentes_campana[nombre_dia_columna].dropna():
        horario_str = str(horario).strip()
        if 'descanso' in horario_str.lower() or '-' not in horario_str:
            continue
        try:
            h_entrada, h_salida = horario_str.split('-')
            h_entrada = normalizar_intervalo(h_entrada)
            h_salida = normalizar_intervalo(h_salida)
            if h_entrada <= intervalo < h_salida:
                agentes_activos += 1
        except:
            continue
    return agentes_activos

@app.route('/api/process', methods=['POST'])
def process_wfm():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
        
    file = request.files['file']
    target_sl = float(request.form.get('target_sl', 80)) / 100.0
    target_time = int(request.form.get('target_time', 20))
    merma = float(request.form.get('merma', 20)) / 100.0
    h_inicio = request.form.get('h_inicio', '09:00')
    h_fin = request.form.get('h_fin', '20:00')
    dias_proyectar = int(request.form.get('dias', 30))

    try:
        xls = pd.ExcelFile(file)
        
        nombre_sheet_llamadas = next((p for p in xls.sheet_names if 'llamada' in p.lower() or 'historico' in p.lower()), xls.sheet_names[0])
        df_llamadas = pd.read_excel(xls, sheet_name=nombre_sheet_llamadas)
        
        nombre_sheet_agentes = next((p for p in xls.sheet_names if 'plantilla' in p.lower() or 'platilla' in p.lower() or 'roster' in p.lower()), None)
        df_agentes = pd.read_excel(xls, sheet_name=nombre_sheet_agentes) if nombre_sheet_agentes else pd.DataFrame()

        df_llamadas.columns = df_llamadas.columns.astype(str).str.strip()
        
        col_recibidas = next((c for c in df_llamadas.columns if any(k in c.lower() for k in ['recibid', 'llamad', 'atendid', 'ofrecid', 'volume'])), None)
        if col_recibidas:
            df_llamadas.rename(columns={col_recibidas: 'Recibidas'}, inplace=True)

        df_llamadas.dropna(subset=['Fecha', 'Intervalo', 'Campaña'], inplace=True)
        df_llamadas['Campaña'] = df_llamadas['Campaña'].astype(str).str.strip()
        df_llamadas['Día'] = df_llamadas['Día'].astype(str).str.strip().str.lower()
        df_llamadas['Intervalo'] = df_llamadas['Intervalo'].apply(normalizar_intervalo)

        reemplazos = {'á':'a', 'é':'e', 'í':'i', 'ó':'o', 'ú':'u'}
        for acento, normal in reemplazos.items():
            df_llamadas['Día'] = df_llamadas['Día'].str.replace(acento, normal)

        if not df_agentes.empty:
            df_agentes.columns = df_agentes.columns.astype(str).str.strip()
            if 'Campaña' in df_agentes.columns:
                df_agentes['Campaña'] = df_agentes['Campaña'].astype(str).str.strip()

        # Obtener todas las campañas únicas
        campanas = df_llamadas['Campaña'].unique().tolist()
        
        df_llamadas['Fecha_Real'] = pd.to_datetime(df_llamadas['Fecha'], dayfirst=True, errors='coerce')
        ultima_fecha = df_llamadas['Fecha_Real'].max()
        fecha_inicio = ultima_fecha + timedelta(days=1)

        mapa_dias = {0: 'Lunes', 1: 'Martes', 2: 'Miércoles', 3: 'Jueves', 4: 'Viernes', 5: 'Sábado', 6: 'Domingo'}
        mapa_dias_lower = {0: 'lunes', 1: 'martes', 2: 'miercoles', 3: 'jueves', 4: 'viernes', 5: 'sabado', 6: 'domingo'}
        
        resultados = []

        # Procesar de forma independiente cada campaña
        for campana in campanas:
            df_camp_data = df_llamadas[df_llamadas['Campaña'] == campana]
            
            cache_semanal = {}
            for i in range(7):
                cache_semanal[mapa_dias_lower[i]] = analizar_dia_historico_campana(df_camp_data, mapa_dias_lower[i])

            for i in range(dias_proyectar):
                fecha_actual = fecha_inicio + timedelta(days=i)
                nombre_dia_col = mapa_dias[fecha_actual.weekday()]
                nombre_dia_hist = mapa_dias_lower[fecha_actual.weekday()]
                
                vol_esperado, curva_v, curva_a = cache_semanal[nombre_dia_hist]
                if vol_esperado == 0: continue
                    
                for intervalo, porcentaje in curva_v.items():
                    intervalo_clean = normalizar_intervalo(intervalo)
                    if not (h_inicio <= intervalo_clean <= h_fin): continue
                        
                    llamadas_exactas = vol_esperado * porcentaje
                    aht = curva_a.get(intervalo, 200) 
                    
                    _, ag_prog_req = calcular_requerimiento(
                        volumen=llamadas_exactas, aht=aht, target_sl=target_sl, target_time=target_time, merma=merma
                    )
                    
                    agentes_reales = contar_agentes_desde_malla(df_agentes, campana, nombre_dia_col, intervalo_clean)
                    delta = agentes_reales - ag_prog_req
                    sl_proyectado = calcular_sl_logrado(llamadas_exactas, aht, target_time, agentes_reales)

                    resultados.append({
                        'Campaña': campana,
                        'Fecha': fecha_actual.strftime('%Y-%m-%d'),
                        'Día_Semana': nombre_dia_col,
                        'Intervalo': intervalo_clean,
                        'Llamadas': int(round(llamadas_exactas)),
                        'AHT': int(round(aht)),
                        'Agentes_Requeridos': int(ag_prog_req),
                        'Agentes_Programados_Reales': int(agentes_reales),
                        'Delta_Net_Staffing': int(delta),
                        'SL_Proyectado': round(sl_proyectado, 1)
                    })

        return jsonify(resultados)

    except Exception as e:
        return jsonify({"error": f"Error procesando Excel: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(port=5000, debug=True)