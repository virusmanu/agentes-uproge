"""Borrador de Ideas Fuerza a partir de las cifras.

Compara el mes evaluado con el mes anterior (y, si está disponible, con igual mes del año anterior).
Solo usa cifras de la Base Consolidada, la Base Glosas y las planillas de Presupuesto Vigente:
las causas, gestiones y datos de gestión (coberturas, beneficiarios, brotes, etc.) quedan como
marcadores [Completar: ...] para el analista. Nunca inventa cifras.
"""
import os, re, glob, calendar, pandas as pd

MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
INST = {
    '01 SUBSE': ('la Subsecretaría de Agricultura', 'Subsecretaría'),
    '02 ODEPA': ('la Oficina de Estudios y Políticas Agrarias (ODEPA)', 'ODEPA'),
    '03 INDAP': ('el Instituto de Desarrollo Agropecuario (INDAP)', 'INDAP'),
    '04 SAG': ('el Servicio Agrícola y Ganadero (SAG)', 'SAG'),
    '05 CONAF': ('la Corporación Nacional Forestal (CONAF)', 'CONAF'),
    '06 CNR': ('la Comisión Nacional de Riego (CNR)', 'CNR'),
}
ORDEN = list(INST)
COMPLETAR = '[Completar por el analista: causas de las variaciones, gestiones en curso y datos de gestión del servicio.]'


# ------------------------------------------------------------------ formato
def mill(m):
    """miles de $ -> '$12.345 millones'"""
    return '$' + f'{m / 1000:,.0f}'.replace(',', '.') + ' millones'

def pct(x):
    return f'{x * 100:,.1f}'.replace('.', ',') + '%'

def pp(x):
    return f'{abs(x) * 100:,.1f}'.replace('.', ',')

def lista(items):
    items = [i for i in items if i]
    if not items: return ''
    return items[0] if len(items) == 1 else ', '.join(items[:-1]) + ' y ' + items[-1]

def sin_codigo(nombre):
    """'24 TRANSFERENCIAS CORRIENTES' -> 'Transferencias Corrientes'; '2401369 Emergencias Agrícolas' -> 'Emergencias Agrícolas'"""
    if not isinstance(nombre, str): return ''
    s = re.sub(r'^\s*\d+\s*', '', nombre).strip()
    if s.isupper():
        menores = {'de', 'del', 'en', 'y', 'a', 'al', 'el', 'la', 'las', 'los', 'por', 'para', 'no', 'con'}
        s = ' '.join(w.capitalize() if (i == 0 or w.lower() not in menores) else w.lower() for i, w in enumerate(s.lower().split()))
    return s

def prog_corto(p):
    """'P02 - Investigación e Innovación ...' -> ('Programa 02', 'Investigación e Innovación ...')"""
    m = re.match(r'P0?(\d+)\s*-\s*(.*)', str(p))
    return (f'Programa {int(m.group(1)):02d}', m.group(2).strip()) if m else (str(p), '')


# ------------------------------------------------------------------ lectura de meses de comparación
def _ent(x):
    """Entero o None (celdas vacías o con texto no numérico no detienen el borrador)."""
    try:
        return None if pd.isna(x) else int(float(x))
    except (TypeError, ValueError):
        return None

def _norm(df):
    df = df.copy()
    for c in ('Subtítulo', 'Ítem', 'Asignación', 'Subasignación'):
        if c in df: df[c] = pd.to_numeric(df[c], errors='coerce').astype('Int64')
    for c in ('Presupuesto Ley', 'Presupuesto Vigente', 'Monto Requerimiento', 'Monto Compromiso', 'Monto Devengo'):
        if c in df: df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
    return df

def consolidado_de(carpeta_mes):
    """Consolidado publicado del mes (si no hay, el generado por el agente)."""
    d = os.path.join(carpeta_mes, 'MINAGRI')
    if not os.path.isdir(d): return None
    fs = [f for f in os.listdir(d) if f.lower().startswith('minagri consolidado a') and f.lower().endswith('.xlsx') and not f.startswith('~$')]
    of = [f for f in fs if 'AGENTE' not in f.upper()]
    elegido = (of or fs or [None])[0]
    return os.path.join(d, elegido) if elegido else None

def leer_consolidado(path):
    return _norm(pd.read_excel(path, sheet_name='Base Consolidada')), _norm(pd.read_excel(path, sheet_name='Base Glosas'))

def carpeta_mes(base, anio, mes):
    if not base or not os.path.isdir(base): return None
    c = [x for x in os.listdir(base) if x.startswith(f'{mes:02d} ') and str(anio) in x and os.path.isdir(os.path.join(base, x))]
    return os.path.join(base, c[0]) if len(c) == 1 else None

def pct_anio_anterior(carpeta_mes_actual, base_anio_ant, anio, mes):
    """% de ejecución de gastos por servicio en igual mes del año anterior.
    1º: 'Cuadro Comparativo' del mes (hoja RESUMEN SERVICIOS); 2º: consolidado del año anterior."""
    d = os.path.join(carpeta_mes_actual, 'MINAGRI')
    cc = [f for f in (os.listdir(d) if os.path.isdir(d) else []) if f.lower().startswith('cuadro comparativo') and f.endswith('.xlsx')]
    if cc:
        try:
            x = pd.read_excel(os.path.join(d, cc[0]), sheet_name='RESUMEN SERVICIOS', header=None)
            fila = next(i for i in range(len(x)) if str(x.iat[i, 1]).strip().upper() == 'INSTITUCIÓN')
            mapa = {'SUBSECRETARIA': '01 SUBSE', 'ODEPA': '02 ODEPA', 'INDAP': '03 INDAP', 'SAG': '04 SAG', 'CONAF': '05 CONAF', 'CNR': '06 CNR', 'TOTAL MINAGRI': 'MINAGRI'}
            out = {}
            for i in range(fila + 1, fila + 10):
                k = mapa.get(str(x.iat[i, 1]).strip().upper())
                if k: out[k] = float(x.iat[i, 8]) / 100
            if len(out) == 7: return out, f'Cuadro Comparativo ({cc[0]})'
        except Exception:
            pass
    c = carpeta_mes(base_anio_ant, anio - 1, mes)
    p = consolidado_de(c) if c else None
    if p:
        b = _norm(pd.read_excel(p, sheet_name='Base Consolidada'))
        g = b[b.Tipo == 'Gastos'].groupby('Servicio')[['Presupuesto Vigente', 'Monto Devengo']].sum()
        out = (g['Monto Devengo'] / g['Presupuesto Vigente']).to_dict()
        out['MINAGRI'] = g['Monto Devengo'].sum() / g['Presupuesto Vigente'].sum()
        return out, os.path.basename(p)
    return {}, None

def desc(d):
    """Descripción del decreto entre comillas, como en las minutas: "Emergencias Sanitarias"."""
    d = (d or '').strip()
    if not d: return ''
    if d.isupper(): d = d.capitalize()
    return '“' + d + '”'


def frase_decreto(n, d, m):
    if abs(m) < 500: return f'N°{n}' + (f' ({desc(d)}, redistribución sin efecto en el total)' if d else ' (redistribución sin efecto en el total)')
    return f'N°{n}' + (f' ({desc(d)}, {"+" if m >= 0 else "−"}{mill(abs(m))})' if d else f' ({"+" if m >= 0 else "−"}{mill(abs(m))})')

def decretos_del_periodo(pvs, desde, hasta):
    """Decretos (Gastos) con fecha en (desde, hasta], por servicio: [(N°, descripción, monto M$)]."""
    out = {}
    for s, p in pvs.items():
        import motor
        b = motor.base_presupuesto(p)          # con las fechas de decreto ya ajustadas al criterio
        b = b[(b.Documento == 'Decreto') & (b.Tipo == 'Gastos') & b['Monto M$'].notna() & (b['Fecha Documento'] > desde) & (b['Fecha Documento'] <= hasta)]
        if len(b):
            b = b.assign(D=b['Descripción Documento'].fillna('').astype(str).str.strip())
            g = b.groupby('Numero Documento').agg(monto=('Monto M$', 'sum'), d=('D', lambda x: '; '.join(dict.fromkeys(v for v in x if v))))
            out[s] = [(_ent(n) if _ent(n) is not None else n, r.d, r.monto) for n, r in g.iterrows()]
    return out


# ------------------------------------------------------------------ análisis
def _agg(b, keys):
    g = b[b.Tipo == 'Gastos'].groupby(keys, dropna=False)[['Presupuesto Vigente', 'Monto Devengo']].sum()
    return g

def tabla(cur, prev, keys):
    a = _agg(cur, keys); a.columns = ['Vigente', 'Devengo']
    if prev is not None:
        p = _agg(prev, keys); p.columns = ['Vigente ant.', 'Devengo ant.']
        a = a.join(p, how='left').fillna(0)
    else:
        a['Vigente ant.'] = 0.0; a['Devengo ant.'] = 0.0
    a['% avance'] = (a['Devengo'] / a['Vigente']).where(a['Vigente'] != 0)
    a['% avance ant.'] = (a['Devengo ant.'] / a['Vigente ant.']).where(a['Vigente ant.'] != 0)
    a['Var. pp'] = a['% avance'] - a['% avance ant.']
    a['Devengo del mes'] = a['Devengo'] - a['Devengo ant.']
    return a

def tabla_glosas(gc, gp):
    k = ['Servicio', 'Glosa']
    a = gc.groupby(k)[['Presupuesto Vigente', 'Monto Devengo']].sum(); a.columns = ['Vigente', 'Devengo']
    if gp is not None:
        p = gp.groupby(k)[['Presupuesto Vigente', 'Monto Devengo']].sum(); p.columns = ['Vigente ant.', 'Devengo ant.']
        a = a.join(p, how='left').fillna(0)
    else:
        a['Vigente ant.'] = 0.0; a['Devengo ant.'] = 0.0
    a['% avance'] = (a['Devengo'] / a['Vigente']).where(a['Vigente'] != 0)
    a['% avance ant.'] = (a['Devengo ant.'] / a['Vigente ant.']).where(a['Vigente ant.'] != 0)
    a['Var. pp'] = a['% avance'] - a['% avance ant.']
    return a


# ------------------------------------------------------------------ redacción
def sobre100(r):
    return ' —gasto superior al presupuesto vigente; [Completar: regularización en trámite]—' if pd.notna(r['% avance']) and r['% avance'] > 1.0005 and r['Vigente'] > 0 else ''

def var_txt(r, sufijo=' en el mes'):
    """'+6,2 pp en el mes' o, si el vigente cambió más de 10% (o es nuevo), la explicación de ese cambio."""
    va, v = r['Vigente ant.'], r['Vigente']
    if not va: return 'sin presupuesto el mes anterior' if v else ''
    if abs(v - va) > 0.10 * abs(va):
        return f'su presupuesto vigente {"aumentó" if v > va else "disminuyó"} {mill(abs(v - va))}{sufijo}'
    if pd.isna(r['Var. pp']): return ''
    return f'{"+" if r["Var. pp"] >= 0 else "−"}{pp(r["Var. pp"])} pp{sufijo}'

def _frase_avance(x, xp, mes_ant, dm, lineal_mes, dv=0, vant=0):
    if xp is None or pd.isna(xp): return ''
    d = x - xp
    ritmo = 'superior' if d > lineal_mes + 0.01 else ('inferior' if d < lineal_mes - 0.01 else 'similar')
    return (f' Respecto de {mes_ant}, la ejecución {"avanzó" if d >= 0 else "retrocedió"} {pp(d)} puntos porcentuales '
            f'({pct(xp)} a {pct(x)}), con un gasto devengado en el mes de {mill(dm)}, ritmo {ritmo} al avance mensual lineal '
            f'de {pp(lineal_mes)} puntos'
            + (f'; la variación del porcentaje está afectada por el aumento de {mill(dv)} del presupuesto vigente.' if vant and dv > 0.05 * vant else '.'))

def _frase_anual(x, xa, anio):
    if xa is None or pd.isna(xa): return ''
    d = x - xa
    if abs(d) < 0.005: return f' Este nivel es similar al {pct(xa)} registrado en igual período de {anio - 1}.'
    return f' Este nivel es {"superior" if d > 0 else "inferior"} en {pp(d)} puntos al {pct(xa)} registrado en igual período de {anio - 1}.'

def _frase_lineal(x, lineal):
    d = x - lineal
    if abs(d) < 0.01: return f' Se ubica en línea con el avance lineal esperado a la fecha ({pct(lineal)}).'
    return f' Se ubica {pp(d)} puntos {"por sobre" if d > 0 else "por debajo de"}l avance lineal esperado a la fecha ({pct(lineal)}).'.replace('por sobrel', 'por sobre el')

def _frase_decretos(s, dec, dv, mes_ant):
    if abs(dv) < 1 and not dec: return ''
    ds = [frase_decreto(n, d, m) for n, d, m in sorted(dec, key=lambda t: -abs(t[2]))]
    if abs(dv) < 1:
        return f' El presupuesto vigente se mantuvo respecto de {mes_ant}; en el período se registraron los Decretos {lista(ds)}, que no alteraron su monto total.' if ds else ''
    base = f' Durante el período, el presupuesto vigente {"aumentó" if dv > 0 else "disminuyó"} en {mill(abs(dv))} respecto de {mes_ant}'
    return base + (f', por efecto de los Decretos {lista(ds)}.' if len(ds) > 1 else f', por efecto del Decreto {ds[0]}.' if ds else '. [Completar: decreto o modificación que explica la variación.]')

def texto_servicio(s, T, TG, anual, dec, mes, anio, mes_ant, lineal, lineal_mes):
    nom, corto = INST[s]
    dia = calendar.monthrange(anio, mes)[1]
    tot = T['serv'].loc[s]
    x, xp = tot['% avance'], tot['% avance ant.']
    p = [f'Al {dia} de {MESES[mes - 1]} de {anio}, {nom} registra un presupuesto vigente de {mill(tot["Vigente"])} y un gasto '
         f'devengado acumulado de {mill(tot["Devengo"])}, equivalente a una ejecución de {pct(x)}.'
         + _frase_avance(x, xp, mes_ant, tot['Devengo del mes'], lineal_mes, tot['Vigente'] - tot['Vigente ant.'], tot['Vigente ant.']) + _frase_anual(x, anual.get(s), anio) + _frase_lineal(x, lineal)
         + _frase_decretos(s, dec.get(s, []), tot['Vigente'] - tot['Vigente ant.'], mes_ant)]
    # subtítulos relevantes
    st = T['subt'].loc[s]
    st = st[st['Vigente'] >= 0.05 * tot['Vigente']].sort_values('Vigente', ascending=False)
    if len(st):
        partes = []
        for (cod, nomb), r in st.iterrows():
            v = var_txt(r)
            partes.append(f'el Subtítulo {int(cod)} {sin_codigo(nomb)} alcanza un {pct(r["% avance"])}{sobre100(r)}, con {mill(r["Devengo"])}' + (f' ({v})' if v else ''))
        f = 'En cuanto a la composición del gasto, ' + lista(partes) + '.'
        top = T['subt'].loc[s].sort_values('Devengo del mes', ascending=False)
        if len(top) and top['Devengo del mes'].iat[0] > 0 and tot['Devengo del mes'] > 0:
            (cod, nomb), r = top.index[0], top.iloc[0]
            f += (f' El mayor gasto del mes se concentró en el Subtítulo {int(cod)} {sin_codigo(nomb)} ({mill(r["Devengo del mes"])}, '
                  f'{pct(r["Devengo del mes"] / tot["Devengo del mes"])} del gasto mensual).')
        p.append(f)
    # programas
    pg = T['prog'].loc[s]
    pg = pg[pg['Vigente'] >= 0.005 * tot['Vigente']]
    if len(pg) > 1:
        partes = []
        for prog, r in pg.sort_values('Vigente', ascending=False).iterrows():
            n, nn = prog_corto(prog); v = var_txt(r, '')
            partes.append(f'{pct(r["% avance"])} en el {n} {nn}' + (f' ({v})' if v else ''))
        bajo = pg.sort_values('% avance').index[0]
        p.append('Por programa, la ejecución alcanza ' + lista(partes) + f'. El menor avance corresponde al {" ".join(prog_corto(bajo))}. [Completar: causa del menor avance.]')
    # asignaciones: mayores pagos del mes y bajo avance
    asg = T['asig'].loc[s]
    asg = asg[[_ent(k[0]) in (24, 31, 32, 33) for k in asg.index]]
    if len(asg):
        may = asg[asg['Devengo del mes'] > max(0.01 * tot['Devengo del mes'], 1)].sort_values('Devengo del mes', ascending=False).head(3)
        if len(may):
            p.append('Los mayores pagos del mes en transferencias, préstamos e inversión corresponden a ' + lista(
                [f'{sin_codigo(k[1])} ({mill(r["Devengo del mes"])}' + (f'; avance acumulado de {pct(r["% avance"])}{sobre100(r)})' if r['Vigente'] > 0 and pd.notna(r['% avance']) else ')') for k, r in may.iterrows()]) + '.')
        rel = asg[(asg['Vigente'] >= 0.01 * tot['Vigente']) & (asg['% avance'] < lineal / 2)].sort_values('Vigente', ascending=False).head(3)
        if len(rel):
            p.append('Presentan bajo avance respecto del lineal esperado ' + lista(
                [f'{sin_codigo(k[1])} ({pct(r["% avance"])} de {mill(r["Vigente"])})' for k, r in rel.iterrows()]) + '. [Completar: estado de convenios, concursos o pagos programados.]')
    # glosas
    if s in TG.index.get_level_values(0):
        g = TG.loc[s].sort_values('Vigente', ascending=False)
        g = g[g['Vigente'] > 0]
        if len(g):
            vt, dt = g['Vigente'].sum(), g['Devengo'].sum()
            alt = g[g['% avance'] >= lineal]
            cero = g[g['Devengo'] <= 0]
            f = f'En materia de glosas presupuestarias, se observa una ejecución global de {pct(dt / vt)} sobre un presupuesto vigente de {mill(vt)}. '
            f += 'Por glosa: ' + lista([f'{gl} {pct(r["% avance"])}' + (f' ({pct(r["% avance ant."])} en {mes_ant})' if pd.notna(r['% avance ant.']) else '') for gl, r in g.iterrows()]) + '.'
            if len(alt):
                f += (f' {"La glosa" if len(alt) == 1 else "Las glosas"} de {lista(list(alt.index))} {"supera" if len(alt) == 1 else "superan"} el avance lineal esperado '
                      f'({pct(lineal)}), por lo que {"requiere" if len(alt) == 1 else "requieren"} seguimiento para no generar déficit al cierre del ejercicio.')
            else:
                f += f' Todas las glosas se mantienen bajo el avance lineal esperado ({pct(lineal)}).'
            if len(cero): f += f' {lista(list(cero.index))} no {"registra" if len(cero) == 1 else "registran"} ejecución a la fecha.'
            p.append(f)
    p.append(COMPLETAR)
    return '\n\n'.join(p)

def texto_minagri(T, TG, anual, dec, mes, anio, mes_ant, lineal, lineal_mes):
    s = T['serv']; tot = s.sum(numeric_only=True)
    x = tot['Devengo'] / tot['Vigente']; xp = tot['Devengo ant.'] / tot['Vigente ant.'] if tot['Vigente ant.'] else None
    p = [f'Al mes de {MESES[mes - 1]} de {anio}, el Ministerio de Agricultura registra una ejecución de {mill(tot["Devengo"])}, equivalente al '
         f'{pct(x)} de su presupuesto vigente de {mill(tot["Vigente"])}.' + _frase_avance(x, xp, mes_ant, tot['Devengo del mes'], lineal_mes)
         + _frase_anual(x, anual.get('MINAGRI'), anio) + _frase_lineal(x, lineal)]
    # servicios
    orden = s.sort_values('% avance', ascending=False)
    p.append('Por servicio, la ejecución alcanza ' + lista([f'{pct(r["% avance"])} en {INST[k][1]}' + (f' ({var_txt(r)})' if var_txt(r) else '')
                                     for k, r in orden.iterrows()]) + '.'
             + f' {INST[orden.index[0]][1]} presenta el mayor avance y {INST[orden.index[-1]][1]} el menor.')
    if anual:
        sup = [INST[k][1] for k in ORDEN if k in anual and s.loc[k, '% avance'] >= anual[k]]
        inf = [(INST[k][1], s.loc[k, '% avance'] - anual[k]) for k in ORDEN if k in anual and s.loc[k, '% avance'] < anual[k]]
        f = 'En comparación con igual período de ' + str(anio - 1) + ', '
        f += (f'{lista(sup)} {"supera" if len(sup) == 1 else "superan"} su avance del año anterior' if sup else 'ningún servicio supera su avance del año anterior')
        if inf:
            inf.sort(key=lambda t: t[1])
            f += '; la diferencia se concentra en ' + lista([f'{n} ({pp(d)} puntos bajo {anio - 1})' for n, d in inf]) + '.'
        else:
            f += '.'
        p.append(f + ' [Completar: factores que explican las diferencias.]')
    # decretos del mes
    todos = [(k, n, d, m) for k, v in dec.items() for n, d, m in v]
    dv = tot['Vigente'] - tot['Vigente ant.']
    if todos or abs(dv) >= 1:
        mayores = sorted(todos, key=lambda t: -abs(t[3]))[:4]
        f = f'Durante el período, el presupuesto vigente sectorial {"aumentó" if dv >= 0 else "disminuyó"} en {mill(abs(dv))}'
        if mayores:
            f += ', destacando ' + lista([f'el Decreto {frase_decreto(n, d, m)} en {INST[k][1]}' for k, n, d, m in mayores])
        p.append(f + '.')
    # glosas en presión
    alt = TG[(TG['Vigente'] > 0) & (TG['% avance'] >= lineal)].sort_values('% avance', ascending=False)
    if len(alt):
        p.append(f'En glosas presupuestarias, superan el avance lineal esperado ({pct(lineal)}) ' + lista([f'{gl.lower()} en {INST[k][1]} ({pct(r["% avance"])})' for (k, gl), r in alt.iterrows()])
                 + ', que requieren gestión para no generar déficit al cierre del ejercicio.')
    else:
        p.append(f'En glosas presupuestarias, todos los servicios se mantienen bajo el avance lineal esperado ({pct(lineal)}).')
    p.append('[Completar por el analista: resultados relevantes por servicio (coberturas, emergencias, temporada de incendios, bonificaciones de riego) y factores que condicionarán el gasto del resto del ejercicio.]')
    return '\n\n'.join(p)


def generar_ideas(base, bglos, prev, prev_glos, anual, dec, mes, anio):
    mes_ant = MESES[mes - 2] if mes > 1 else None
    lineal = mes / 12; lineal_mes = 1 / 12
    base = _norm(base); prev = _norm(prev) if prev is not None else None
    con_subt = lambda b: b[b['Subtítulo'].notna()] if b is not None else None
    con_asig = lambda b: b[b['Nombre Asignación'].notna() & b['Subtítulo'].notna()] if b is not None else None
    T = {'serv': tabla(base, prev, ['Servicio']),
         'subt': tabla(con_subt(base), con_subt(prev), ['Servicio', 'Subtítulo', 'Nombre Subtítulo']),
         'prog': tabla(base, prev, ['Servicio', 'Programa']),
         'asig': tabla(con_asig(base), con_asig(prev), ['Servicio', 'Subtítulo', 'Nombre Asignación'])}
    TG = tabla_glosas(bglos, prev_glos)
    filas = [(0, 'MINAGRI', texto_minagri(T, TG, anual, dec, mes, anio, mes_ant, lineal, lineal_mes))]
    for i, s in enumerate(ORDEN, 1):
        filas.append((i, s.split()[1], texto_servicio(s, T, TG, anual, dec, mes, anio, mes_ant, lineal, lineal_mes)))
    ideas = pd.DataFrame(filas, columns=['Código Institución', 'Institución', 'Ideas Fuerza'])
    # hoja de respaldo con las cifras usadas
    resp = []
    for nivel, k in (('Servicio', 'serv'), ('Subtítulo', 'subt'), ('Programa', 'prog')):
        t = T[k].reset_index(); t.insert(0, 'Nivel', nivel); resp.append(t)
    g = TG.reset_index(); g.insert(0, 'Nivel', 'Glosa'); resp.append(g)
    respaldo = pd.concat(resp, ignore_index=True)
    cols = ['Nivel', 'Servicio', 'Subtítulo', 'Nombre Subtítulo', 'Programa', 'Glosa', 'Vigente', 'Devengo', '% avance', 'Vigente ant.', 'Devengo ant.', '% avance ant.', 'Var. pp', 'Devengo del mes']
    respaldo = respaldo[[c for c in cols if c in respaldo.columns]]
    respaldo['% año anterior'] = respaldo.apply(lambda r: anual.get(r['Servicio']) if r['Nivel'] == 'Servicio' else None, axis=1)
    return ideas, respaldo
