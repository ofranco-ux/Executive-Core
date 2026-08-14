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
    req_arr = np.array(req_vector, dtype=float)

    cob_efectiva = np.zeros(m, dtype=float)
    x_turnos_dict = {}

    agentes_nocturnos_totales = 0
    agentes_diurnos_totales = 0

    # --- PASO 1: TURNO NOCTURNO FIJO (22:00 A 07:00 / 5x2) ---
    if es_nocturno:
        label_jornada_noc = "9.0 hrs (Nocturno 5x2)"
        reqs_nocturnos = []
        indices_nocturnos = []

        for j in range(m):
            min_in = parse_time_str(intervalos[j])
            if min_in is not None:
                # Ventana nocturna: 22:00 a 07:00
                if min_in >= (22 * 60) or min_in < (7 * 60):
                    reqs_nocturnos.append(req_arr[j])
                    indices_nocturnos.append(j)

        # En lugar de cubrir el pico absoluto, se calcula el requerimiento promedio ajustado por asistencia
        promedio_req_noc = np.mean(reqs_nocturnos) if reqs_nocturnos else 0
        if promedio_req_noc > 0:
            # Agentes en plantilla requeridos para cubrir el turno nocturno
            agentes_nocturnos = max(1, int(round(promedio_req_noc)))
            key_turno_noc = ("22:00", "07:00", label_jornada_noc)
            x_turnos_dict[key_turno_noc] = agentes_nocturnos
            agentes_nocturnos_totales = agentes_nocturnos

            # La cobertura EFECTIVA en mesa considera el porcentaje real asistido (descontando shrinkage)
            cob_real_noc = agentes_nocturnos * factor_asistencia
            for idx in indices_nocturnos:
                cob_efectiva[idx] = cob_real_noc

    # --- PASO 2: TURNOS DIURNOS Y AJUSTE INTRADÍA ---
    inicio_global, fin_global = obtener_ventana_global(campanas_activas)
    duracion_jornada = float(duracion_jornada)
    SHIFT_BLOCKS = int(round(duracion_jornada * 2))
    duracion_minutos = int(round(duracion_jornada * 60))
    label_jornada_diurna = f"{duracion_jornada:.1f} hrs".replace('.0', '')

    for j in range(m):
        min_in = parse_time_str(intervalos[j])
        if min_in is None:
            continue

        # Evaluar únicamente intervalos diurnos (entre apertura y cierre global)
        if min_in >= inicio_global and min_in < fin_global:
            deficit = req_arr[j] - cob_efectiva[j]
            if deficit > 0.2:
                # Asignar agentes considerando merma para cubrir solo el déficit real
                agentes_necesarios = max(1, int(round(deficit / factor_asistencia)))
                min_out = (min_in + duracion_minutos) % (24 * 60)
                h_out = f"{(int(min_out // 60)):02d}:{(int(min_out % 60)):02d}"
                
                key_turno = (intervalos[j], h_out, label_jornada_diurna)
                x_turnos_dict[key_turno] = x_turnos_dict.get(key_turno, 0) + agentes_necesarios

                # Aplicar la cobertura efectiva para la duración de este turno
                for t in range(j, min(j + SHIFT_BLOCKS, m)):
                    min_t = parse_time_str(intervalos[t])
                    # No extender cobertura diurna dentro del rango nocturno estricto
                    if min_t is not None and (min_t < (22 * 60) and min_t >= (7 * 60)):
                        cob_efectiva[t] += agentes_necesarios * factor_asistencia

    # --- PASO 3: CÁLCULO DE METRICAS PROYECTADAS (SL & STAFFING) ---
    sl_optimo_vector = []
    for i in range(m):
        c = llamadas_arr[i]
        aht_s = aht_arr[i]
        n_opt = cob_efectiva[i]
        a_erl = (c * aht_s) / 1800.0 if (c > 0 and aht_s > 0) else 0.0
        
        # Calcular SL con agentes efectivos reales en mesa
        sl_val = erlang_c_sl_optimizado(a_erl, n_opt, aht_s, target_time) if c > 0 else 100.0
        sl_optimo_vector.append(sl_val)

    sl_arr = np.array(sl_optimo_vector)
    sl_optimo_global = float(np.sum(llamadas_arr * sl_arr) / tot_llamadas) if tot_llamadas > 0 else 100.0
    sl_optimo_global = round(sl_optimo_global, 1)

    cobertura_entera = np.round(cob_efectiva).astype(int).tolist()
    turnos_sugeridos = []
    total_agentes_diarios = 0

    for (h_in, h_out, label_dur), qty in x_turnos_dict.items():
        if qty > 0:
            turnos_sugeridos.append({
                'horario_entrada': h_in,
                'horario_salida': h_out,
                'agentes_a_programar': qty,
                'duracion': label_dur
            })
            total_agentes_diarios += qty
            if "Nocturno" not in label_dur:
                agentes_diurnos_totales += qty

    # Headcount semanal ponderado
    hc_nocturno = math.ceil(agentes_nocturnos_totales * (7.0 / 5.0))
    hc_diurno = math.ceil(agentes_diurnos_totales * (7.0 / 6.0))
    headcount_semanal_requerido = hc_nocturno + hc_diurno

    total_req = np.sum(req_arr)
    total_prog_efec = np.sum(cob_efectiva)
    staffing_level_optimo = round(float((total_prog_efec / total_req * 100.0)), 1) if total_req > 0 else 100.0
    eficiencia = round(min(100.0, (total_req / total_prog_efec * 100.0)), 1) if total_prog_efec > 0 else 100.0

    return turnos_sugeridos, cobertura_entera, total_agentes_diarios, headcount_semanal_requerido, eficiencia, sl_optimo_vector, sl_optimo_global, staffing_level_optimo
