"""Revisión del consolidado del mes: controles con semáforo, comentarios del analista y hoja de control.

Todo se calcula sobre el DEVENGO y el presupuesto vigente. No modifica el consolidado que lee Power BI, el Catálogo
Maestro ni ninguna planilla: escribe solo archivos propios en la carpeta MINAGRI del mes
  - MINAGRI Revisión a <Mes> <Año>.xlsx       (alertas, estado, comentarios)
  - MINAGRI Hoja de control a <Mes> <Año>.xlsx (resumen para el visto bueno de la jefatura)
Los umbrales y los pares de transferencias intrapartida están en "Reglas de revisión.xlsx", junto al programa."""
import os, datetime, getpass
import pandas as pd

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REGLAS = os.path.join(RAIZ, 'Reglas de revisión.xlsx')
MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
ROJO, AMARILLO = 'Rojo', 'Amarillo'
COLS = ['Nivel', 'Control', 'Servicio', 'Programa', 'Código', 'Nombre', 'Detalle', 'Vigente', 'Devengo', 'Referencia',
        'Estado', 'Comentario', 'Usuario', 'Fecha', 'Clave']

UMBRALES = [  # (clave, valor, descripción)
    ('cuadratura_m$', 1, 'Cuadratura: diferencia máxima (miles de $) entre el devengo del informe bruto y el consolidado'),
    ('sobre_vigente_rojo', 1.00, 'Devengo sobre vigente: rojo si el avance supera este valor (1 = 100%)'),
    ('sobre_vigente_amarillo', 0.95, 'Devengo sobre vigente: amarillo desde este avance, solo antes del mes indicado'),
    ('sobre_vigente_amarillo_hasta_mes', 10, 'Último mes en que se aplica el amarillo de devengo sobre vigente (10 = octubre)'),
    ('retroceso_m$', 1, 'Devengo acumulado menor que el del mes anterior: diferencia mínima (miles de $)'),
    ('salto_veces', 3, 'Salto del mes: devengo del mes mayor que N veces el promedio mensual anterior'),
    ('sin_ejecucion_desde_mes', 6, 'Sin ejecución: se revisa desde este mes (6 = junio)'),
    ('materialidad', 0.005, 'Asignación relevante (para los amarillos): vigente igual o mayor a esta proporción del vigente del servicio (0,005 = 0,5%)'),
    ('glosa_rojo', 1.00, 'Glosas: rojo si el devengo supera este avance del límite'),
    ('glosa_amarillo', 0.90, 'Glosas: amarillo desde este avance del límite'),
    ('intrapartida_m$', 1, 'Transferencias intrapartida: diferencia máxima (miles de $) entre lo que transfiere SUBSE y lo que recibe el servicio'),
]
EXCEDIBLES = [  # (código o prefijo, denominación, fundamento, activo)
    ('2599', 'Otros Íntegros al Fisco', 'Ley de Presupuestos 2026, art. 34, en relación con el art. 28 del DL N°1.263: se puede exceder y se sanciona después por decreto de Hacienda (art. 70 DL 1.263)', 'Sí'),
    ('3407', 'Deuda Flotante', 'Ley de Presupuestos 2026, art. 34, en relación con el art. 28 del DL N°1.263: se puede exceder y se sanciona después por decreto de Hacienda (art. 70 DL 1.263)', 'Sí'),
    ('2303', 'Prestaciones Sociales del Empleador', 'Art. 28 del DL N°1.263 (gasto excedible) – FUNDAMENTO A CONFIRMAR POR UPROGE', 'Sí'),
    ('2602', 'Compensaciones por Daños a Terceros y/o a la Propiedad', 'Art. 28 del DL N°1.263 (gasto excedible) – FUNDAMENTO A CONFIRMAR POR UPROGE', 'Sí'),
]
CRITERIOS = [  # (clave, valor, descripción)
    ('criterio_fecha_decreto', 'Toma de razón', 'Fecha desde la que un decreto de modificación se aplica al presupuesto vigente: "Toma de razón" (criterio UPROGE, '
     'se verifica en DIPRES), "Decreto" (fecha del decreto en DIPRES) o "Planilla" (la que registra cada servicio)'),
]
PARES = [  # Subsecretaría (Gastos, P04 BID)  ->  servicio receptor (Ingresos)
    ('01 SUBSE', 2402301, '03 INDAP', 'P01', 502032),
    ('01 SUBSE', 3302130, '03 INDAP', 'P01', 1302030),
    ('01 SUBSE', 2402302, '04 SAG', 'P01', 502024),
    ('01 SUBSE', 2402303, '04 SAG', 'P07', 502025),
    ('01 SUBSE', 2402304, '04 SAG', 'P09', 502026),
    ('01 SUBSE', 3302131, '04 SAG', 'P01', 1302201),
    ('01 SUBSE', 3302132, '04 SAG', 'P07', 1302202),
    ('01 SUBSE', 3302133, '04 SAG', 'P09', 1302203),
]


# ------------------------------------------------------------------ reglas
def leer_reglas():
    """Umbrales y pares intrapartida. Si el archivo no existe, se crea con los valores acordados por UPROGE."""
    if not os.path.exists(REGLAS):
        crear_reglas()
    try:
        u = pd.read_excel(REGLAS, sheet_name='Umbrales')
        U = {str(k).strip(): float(v) for k, v in zip(u['Clave'], u['Valor']) if pd.notna(k) and pd.notna(v)}
    except Exception:
        U = {}
    for k, v, _ in UMBRALES: U.setdefault(k, v)
    try:
        e = pd.read_excel(REGLAS, sheet_name='Excedibles', dtype=str).fillna('')
    except Exception:
        e = None
    if e is None:                                   # archivo de una versión anterior: se agrega la hoja con los valores iniciales
        try: _agregar_hoja_excedibles()
        except Exception: pass
        X = [(c, n, f) for c, n, f, a in EXCEDIBLES if a == 'Sí']
    else:
        X = [(str(r['Código']).strip(), r['Denominación'], r['Fundamento']) for _, r in e.iterrows()
             if str(r['Código']).strip() and str(r['Activo']).strip().lower() in ('sí', 'si', 's', 'x')]
    U['_excedibles'] = X
    try:
        c = pd.read_excel(REGLAS, sheet_name='Criterios', dtype=str).fillna('')
        C = {str(k).strip(): str(v).strip() for k, v in zip(c['Clave'], c['Valor'])}
    except Exception:
        try: _agregar_hoja_criterios()
        except Exception: pass
        C = {}
    for k, v, _ in CRITERIOS: C[k] = C.get(k) or v
    U['_criterios'] = C
    try:
        p = pd.read_excel(REGLAS, sheet_name='Transferencias intrapartida')
        P = [(r['Servicio que transfiere'], int(r['Asignación de gasto']), r['Servicio que recibe'], str(r['Programa que recibe']).strip(), int(r['Asignación de ingreso']))
             for _, r in p.iterrows() if pd.notna(r['Asignación de gasto']) and pd.notna(r['Asignación de ingreso'])]
    except Exception:
        P = list(PARES)
    return U, P

def crear_reglas():
    from openpyxl.styles import Font, PatternFill
    with pd.ExcelWriter(REGLAS, engine='openpyxl') as w:
        pd.DataFrame(UMBRALES, columns=['Clave', 'Valor', 'Descripción']).to_excel(w, sheet_name='Umbrales', index=False)
        pd.DataFrame(EXCEDIBLES, columns=['Código', 'Denominación', 'Fundamento', 'Activo']).to_excel(w, sheet_name='Excedibles', index=False)
        pd.DataFrame(CRITERIOS, columns=['Clave', 'Valor', 'Descripción']).to_excel(w, sheet_name='Criterios', index=False)
        pd.DataFrame(PARES, columns=['Servicio que transfiere', 'Asignación de gasto', 'Servicio que recibe', 'Programa que recibe', 'Asignación de ingreso']
                     ).to_excel(w, sheet_name='Transferencias intrapartida', index=False)
        for ws in w.book.worksheets:
            for c in ws[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E5F')
            for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = min(max(len(str(c.value or '')) for c in col) + 2, 110)
            for c in ws['B'][1:]: c.fill = PatternFill('solid', fgColor='FFF2CC')


def _agregar_hoja_excedibles():
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    wb = openpyxl.load_workbook(REGLAS)
    if 'Excedibles' in wb.sheetnames: return
    ws = wb.create_sheet('Excedibles', 1)
    ws.append(['Código', 'Denominación', 'Fundamento', 'Activo'])
    for r in EXCEDIBLES: ws.append(list(r))
    for c in ws[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E5F')
    for col, w in zip('ABCD', (10, 50, 120, 8)): ws.column_dimensions[col].width = w
    wb.save(REGLAS)

def _agregar_hoja_criterios():
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    wb = openpyxl.load_workbook(REGLAS)
    if 'Criterios' in wb.sheetnames: return
    ws = wb.create_sheet('Criterios', 2)
    ws.append(['Clave', 'Valor', 'Descripción'])
    for r in CRITERIOS: ws.append(list(r))
    for c in ws[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E5F')
    for c in ws['B'][1:]: c.fill = PatternFill('solid', fgColor='FFF2CC')
    for col, w in zip('ABC', (26, 16, 150)): ws.column_dimensions[col].width = w
    wb.save(REGLAS)

def criterio_fecha_decreto():
    try: return leer_reglas()[0]['_criterios']['criterio_fecha_decreto']
    except Exception: return CRITERIOS[0][1]

def excedible(cod, X):
    """(denominación, fundamento) si el código (subtítulo, ítem o asignación) está dentro de un concepto excedible."""
    c = '' if cod is None or (isinstance(cod, float) and pd.isna(cod)) else str(int(cod)) if not isinstance(cod, str) else cod
    for pre, nom, fun in X:
        if c.startswith(pre):
            return nom, fun
    return None


# ------------------------------------------------------------------ utilidades
def _num(x):
    return pd.to_numeric(x, errors='coerce')

def _niveles(df):
    """Clave de comparación: asignación si el vigente está a ese nivel; si no, ítem; si no, subtítulo."""
    d = df.copy()
    for c in ('Subtítulo', 'Ítem', 'Asignación'): d[c] = _num(d[c])
    vig = d[d['Presupuesto Vigente'] != 0]
    a_nivel_item = set(map(tuple, vig[vig['Ítem'].notna() & vig['Asignación'].isna()][['Servicio', 'Programa', 'Ítem']].dropna().values))
    a_nivel_subt = set(map(tuple, vig[vig['Ítem'].isna()][['Servicio', 'Programa', 'Subtítulo']].dropna().values))
    cod, nom = [], []
    for s, p, st, it, asg, nst, nit, nas in zip(d.Servicio, d.Programa, d['Subtítulo'], d['Ítem'], d['Asignación'],
                                                d['Nombre Subtítulo'], d['Nombre Ítem'], d['Nombre Asignación']):
        if (s, p, st) in a_nivel_subt or pd.isna(it):
            cod.append(st); nom.append(nst)
        elif (s, p, it) in a_nivel_item or pd.isna(asg):
            cod.append(it); nom.append(nit)
        else:
            cod.append(asg); nom.append(nas)
    d['_cod'] = cod; d['_nom'] = nom
    return d

def _tabla(df):
    d = _niveles(df[df.Tipo == 'Gastos'])
    d = d[d['_cod'].notna()]
    g = d.groupby(['Servicio', 'Programa', '_cod'])[['Presupuesto Vigente', 'Monto Devengo']].sum()
    n = d.groupby(['Servicio', 'Programa', '_cod'])['_nom'].first()
    return g.join(n)

def _alerta(nivel, control, s, prog, cod, nom, detalle, vig=None, dev=None, ref=None):
    return {'Nivel': nivel, 'Control': control, 'Servicio': s, 'Programa': prog, 'Código': '' if cod is None or pd.isna(cod) else str(int(cod)) if not isinstance(cod, str) else cod,
            'Nombre': '' if nom is None or (isinstance(nom, float) and pd.isna(nom)) else str(nom), 'Detalle': detalle,
            'Vigente': vig, 'Devengo': dev, 'Referencia': ref, 'Estado': 'Pendiente', 'Comentario': '', 'Usuario': '', 'Fecha': ''}

def _m(v):
    return f'{v:,.0f}'.replace(',', '.')

def _p(v):
    return f'{v * 100:,.1f}'.replace(',', 'X').replace('.', ',').replace('X', '.') + '%'


# ------------------------------------------------------------------ controles
def _control_fechas_dipres(corte, crit):
    """Contrasta la fecha que registra cada planilla con las fechas DIPRES (verificadas en decretos.verificar)."""
    try:
        import decretos as DC
        info = (DC.ULTIMO or {}).get('fechas', {})
    except Exception:
        info = {}
    A = []; usa_decreto = {}; c = pd.Timestamp(corte) if corte is not None else None
    for n, i in sorted(info.items()):
        ap = i.get('aplicada'); tr = i.get('manual') or i.get('tr'); fd = i.get('decreto')
        for s, fp in i['planillas'].items():
            if fp is None or pd.isna(fp): continue
            fp = pd.Timestamp(fp)
            if ap is None: continue      # no identificado en DIPRES: se usa la fecha de la planilla (normal, ver hoja Decretos)
            if fp == pd.Timestamp(ap): continue
            if crit == 'Toma de razón' and fd is not None and fp == pd.Timestamp(fd):
                usa_decreto.setdefault(s, []).append((n, fp, pd.Timestamp(ap)))
                continue
            cruza = c is not None and ((fp <= c) != (pd.Timestamp(ap) <= c))
            A.append(_alerta(ROJO if cruza else AMARILLO, 'Fecha de decreto no coincide con DIPRES', s, '', f'N°{n}', 'Decreto',
                             f'La planilla registra el decreto N°{n} con fecha {fp:%d-%m-%Y}; en DIPRES el decreto es del '
                             + (f'{pd.Timestamp(fd):%d-%m-%Y}' if fd is not None else '[sin fecha]')
                             + (f' y la toma de razón del {pd.Timestamp(tr):%d-%m-%Y}' if tr is not None else '')
                             + f'. El agente aplicó {pd.Timestamp(ap):%d-%m-%Y}'
                             + (f'; con la fecha de la planilla el decreto habría quedado {"dentro" if fp <= c else "fuera"} del corte del {c:%d-%m-%Y}. '
                                'Revise que el N° de decreto esté bien digitado en la planilla' if cruza else '. Revise la planilla del servicio')))
    for s, l in usa_decreto.items():
        cr = [x for x in l if c is not None and ((x[1] <= c) != (x[2] <= c))]
        A.append(_alerta(AMARILLO, 'Planilla registra fecha del decreto (no la toma de razón)', s, '', '', 'Decretos',
                         f'{len(l)} decreto(s) registrados con la fecha del decreto en vez de la toma de razón: '
                         + ', '.join(f'N°{n} ({fp:%d-%m-%Y} → TR {ap:%d-%m-%Y})' for n, fp, ap in l)
                         + ('. El agente aplicó la toma de razón' + (f'; con el corte del {c:%d-%m-%Y} esto cambia el vigente de: '
                            + ', '.join(f'N°{n}' for n, *_ in cr) if cr else '') + '. Conviene unificar el criterio con el servicio')))
    return A

def controles(base, bg, prev, mes, cuadratura, pvs, U, P, corte=None):
    A = []
    # 1. Cuadratura del devengo: informe bruto = consolidado + excluido por programa

    cons = base.groupby(['Servicio', 'Tipo'])['Monto Devengo'].sum()
    for (s, t), c in sorted(cuadratura.items()):
        dif = c['Devengo informe'] - c['Devengo excluido por programa'] - float(cons.get((s, t), 0))
        if abs(dif) > U['cuadratura_m$']:
            A.append(_alerta(ROJO, 'Cuadratura del devengo', s, '', None, t,
                             f'El informe bruto suma {_m(c["Devengo informe"])} M$ de devengo; consolidado {_m(cons.get((s, t), 0))} M$ y excluido por programa '
                             f'{_m(c["Devengo excluido por programa"])} M$ (diferencia {_m(dif)} M$)', None, c['Devengo informe'], dif))
    T = _tabla(base)
    vig_serv = base[base.Tipo == 'Gastos'].groupby('Servicio')['Presupuesto Vigente'].sum()
    relevante = lambda s, v: v >= U['materialidad'] * vig_serv.get(s, 0) and v > 0
    # 2. Devengo sobre vigente
    for (s, prog, cod), r in T.iterrows():
        v, d = r['Presupuesto Vigente'], r['Monto Devengo']
        if d <= 0: continue
        ex = excedible(cod, U['_excedibles'])
        if ex and d > v * U['sobre_vigente_rojo'] + U['cuadratura_m$']:
            A.append(_alerta(AMARILLO, 'Excedible por regularizar', s, prog, cod, r['_nom'],
                             f'Gasto excedible: devengo {_m(d)} M$ sobre vigente {_m(v)} M$ (exceso {_m(d - v)} M$, a sancionar por decreto). {ex[1]}', v, d, d - v))
            continue
        if v <= 0:
            A.append(_alerta(ROJO, 'Devengo sin presupuesto vigente', s, prog, cod, r['_nom'], f'Devengo {_m(d)} M$ sin presupuesto vigente', v, d))
        elif d > v * U['sobre_vigente_rojo'] + U['cuadratura_m$']:   # tolerancia de redondeo
            A.append(_alerta(ROJO, 'Devengo sobre el vigente', s, prog, cod, r['_nom'], f'Avance {_p(d / v)}: devengo {_m(d)} M$ sobre vigente {_m(v)} M$', v, d, d / v))
        elif d / v >= U['sobre_vigente_amarillo'] and mes <= U['sobre_vigente_amarillo_hasta_mes'] and relevante(s, v):
            A.append(_alerta(AMARILLO, 'Devengo cerca del vigente', s, prog, cod, r['_nom'], f'Avance {_p(d / v)} en {MESES[mes - 1].lower()}', v, d, d / v))
    # 3 y 4. Comparación con el mes anterior
    if prev is not None and mes > 1:
        Tp = _tabla(prev)
        J = T.join(Tp[['Monto Devengo']].rename(columns={'Monto Devengo': 'Dev ant'}), how='left')
        for (s, prog, cod), r in J.iterrows():
            d, da, v = r['Monto Devengo'], r['Dev ant'], r['Presupuesto Vigente']
            if pd.isna(da): continue
            if d < da - U['retroceso_m$']:
                A.append(_alerta(ROJO, 'Devengo menor que el mes anterior', s, prog, cod, r['_nom'],
                                 f'Devengo acumulado {_m(d)} M$; en {MESES[mes - 2].lower()} era {_m(da)} M$ (baja {_m(da - d)} M$)', v, d, da))
            elif mes > 2 and relevante(s, v):
                mensual, prom = d - da, da / (mes - 1)
                if prom > 0 and mensual > U['salto_veces'] * prom:
                    A.append(_alerta(AMARILLO, 'Salto del mes', s, prog, cod, r['_nom'],
                                     f'Devengo del mes {_m(mensual)} M$, {mensual / prom:,.1f}'.replace('.', ',') + f' veces el promedio mensual anterior ({_m(prom)} M$)', v, d, mensual))
    # 5. Sin ejecución
    if mes >= U['sin_ejecucion_desde_mes']:
        for (s, prog, cod), r in T.iterrows():
            if relevante(s, r['Presupuesto Vigente']) and r['Monto Devengo'] == 0:
                A.append(_alerta(AMARILLO, 'Sin ejecución', s, prog, cod, r['_nom'], f'Vigente {_m(r["Presupuesto Vigente"])} M$ sin devengo a {MESES[mes - 1].lower()}',
                                 r['Presupuesto Vigente'], 0.0))
    # 6. Glosas
    G = bg.groupby(['Servicio', 'Programa', 'Glosa'])[['Presupuesto Vigente', 'Monto Devengo']].sum()
    for (s, prog, gl), r in G.iterrows():
        v, d = r['Presupuesto Vigente'], r['Monto Devengo']
        if v <= 0 or d <= 0: continue
        if d > v * U['glosa_rojo'] + U['cuadratura_m$']:
            A.append(_alerta(ROJO, 'Glosa sobre su límite', s, prog, None, gl, f'Devengo {_m(d)} M$ sobre el límite vigente {_m(v)} M$ ({_p(d / v)})', v, d, d / v))
        elif d / v >= U['glosa_amarillo']:
            A.append(_alerta(AMARILLO, 'Glosa cerca de su límite', s, prog, None, gl, f'Avance {_p(d / v)} del límite vigente', v, d, d / v))
    # 7. Decretos registrados con fechas distintas entre planillas (solo si se usa la fecha de cada planilla;
    #    con el criterio de toma de razón el agente aplica una sola fecha DIPRES a todos los servicios)
    crit = U.get('_criterios', {}).get('criterio_fecha_decreto', 'Toma de razón')
    if pvs and crit == 'Planilla':
        filas = []
        for s, ruta in pvs.items():
            try:
                b = pd.read_excel(ruta, sheet_name='Base Presupuesto', usecols=['Documento', 'Numero Documento', 'Fecha Documento'])
            except Exception:
                continue
            b = b[(b.Documento == 'Decreto') & b['Numero Documento'].notna() & b['Fecha Documento'].notna()]
            for n, f in b.groupby('Numero Documento')['Fecha Documento'].min().items():
                filas.append((int(n), s, pd.Timestamp(f)))
        R = pd.DataFrame(filas, columns=['N', 'Servicio', 'Fecha'])
        for n, g in R.groupby('N'):
            # solo importa si la fecha de corte del mes deja el decreto dentro en una planilla y fuera en otra
            if g.Fecha.nunique() > 1 and (corte is None or (g.Fecha.min() <= pd.Timestamp(corte) < g.Fecha.max())):
                det = '; '.join(f'{s} {f:%d-%m-%Y}' for s, f in zip(g.Servicio, g.Fecha))
                A.append(_alerta(ROJO, 'Decreto con fechas distintas entre planillas', ', '.join(g.Servicio), '', f'N°{n}', 'Decreto',
                                 f'El decreto N°{n} está registrado con fechas distintas en las planillas de Presupuesto Vigente ({det}); '
                                 + (f'con el corte del {pd.Timestamp(corte):%d-%m-%Y} quedó incluido en unos servicios y excluido en otros' if corte is not None else
                                    'con la fecha de corte puede quedar incluido en un servicio y excluido en otro')))
    # 9. Fechas de las planillas contrastadas con DIPRES (decreto y toma de razón)
    if crit != 'Planilla':
        A.extend(_control_fechas_dipres(corte, crit))
    # 8. Transferencias intrapartida: lo que transfiere SUBSE = lo que recibe el servicio
    def vig(s, tipo, cod, prog=None):
        x = base[(base.Servicio == s) & (base.Tipo == tipo) & (_num(base['Asignación']) == cod)]
        if prog: x = x[x.Programa.astype(str).str.startswith(prog + ' ')]
        return float(x['Presupuesto Vigente'].sum()), x
    for s1, a1, s2, p2, a2 in P:
        v1, x1 = vig(s1, 'Gastos', a1); v2, x2 = vig(s2, 'Ingresos', a2, p2)
        if len(x1) == 0 and len(x2) == 0: continue
        if abs(v1 - v2) > U['intrapartida_m$']:
            n1 = x1['Nombre Asignación'].iloc[0] if len(x1) else str(a1)
            A.append(_alerta(ROJO, 'Transferencia intrapartida descuadrada', f'{s1} → {s2}', p2, a1, n1,
                             f'{s1} transfiere {_m(v1)} M$ (gasto {a1}) y {s2} {p2} recibe {_m(v2)} M$ (ingreso {a2}); diferencia {_m(v1 - v2)} M$. '
                             'Revise que el decreto que modificó la transferencia esté en ambas planillas con la misma fecha', v1, None, v2))
    df = pd.DataFrame(A, columns=[c for c in COLS if c != 'Clave'])
    df['Clave'] = df.Control + '|' + df.Servicio.astype(str) + '|' + df.Programa.astype(str) + '|' + df['Código'].astype(str) + '|' + df.Nombre.astype(str)
    orden = {ROJO: 0, AMARILLO: 1}
    return df.sort_values(['Nivel', 'Servicio', 'Control'], key=lambda c: c.map(orden) if c.name == 'Nivel' else c).reset_index(drop=True)


# ------------------------------------------------------------------ archivo de revisión
def ruta_revision(carpeta_mes, anio, mes):
    return os.path.join(carpeta_mes, 'MINAGRI', f'MINAGRI Revisión a {MESES[mes - 1]} {anio}.xlsx')

def ruta_hoja_control(carpeta_mes, anio, mes):
    return os.path.join(carpeta_mes, 'MINAGRI', f'MINAGRI Hoja de control a {MESES[mes - 1]} {anio}.xlsx')

def _leer(ruta):
    try:
        return pd.read_excel(ruta, sheet_name='Alertas', dtype={'Comentario': str, 'Usuario': str, 'Fecha': str, 'Clave': str}).fillna('')
    except Exception:
        return None

def revisar(base, bg, prev, anio, mes, carpeta_mes, carpeta_mes_anterior, cuadratura, pvs, corte):
    """Corre los controles, recupera comentarios de esta revisión y aceptaciones del mes anterior, y guarda el archivo."""
    U, P = leer_reglas()
    df = controles(base, bg, prev, mes, cuadratura, pvs, U, P, corte)
    ant = _leer(ruta_revision(carpeta_mes, anio, mes))                 # misma revisión, generada antes
    if ant is not None and len(ant):
        k = {r.Clave: r for r in ant.itertuples() if str(r.Estado).startswith('Aceptada')}
        for i, c in df.Clave.items():
            if c in k: df.loc[i, ['Estado', 'Comentario', 'Usuario', 'Fecha']] = [k[c].Estado, k[c].Comentario, k[c].Usuario, k[c].Fecha]
    if carpeta_mes_anterior and mes > 1:                                # aceptadas el mes anterior
        pa = _leer(ruta_revision(carpeta_mes_anterior, anio, mes - 1))
        if pa is not None and len(pa):
            k = {r.Clave: r for r in pa.itertuples() if str(r.Estado).startswith('Aceptada')}
            for i, c in df.Clave.items():
                if df.loc[i, 'Estado'] == 'Pendiente' and c in k:
                    df.loc[i, ['Estado', 'Comentario', 'Usuario', 'Fecha']] = [f'Aceptada en {MESES[mes - 2].lower()}', k[c].Comentario, k[c].Usuario, k[c].Fecha]
    guardar(df, carpeta_mes, anio, mes, corte)
    return df

def resumen(df):
    pend = df[df.Estado == 'Pendiente']
    return {'rojos': int((df.Nivel == ROJO).sum()), 'amarillos': int((df.Nivel == AMARILLO).sum()),
            'rojos_pendientes': int((pend.Nivel == ROJO).sum()), 'amarillos_pendientes': int((pend.Nivel == AMARILLO).sum()),
            'aceptadas': int(df.Estado.str.startswith('Aceptada').sum())}

def guardar(df, carpeta_mes, anio, mes, corte):
    from openpyxl.styles import Font, PatternFill, Alignment
    ruta = ruta_revision(carpeta_mes, anio, mes); r = resumen(df)
    dec = None                      # trazabilidad de las fechas de decretos del mes (hoja Decretos)
    try:
        import decretos as DC
        if (DC.ULTIMO or {}).get('anio') == anio and DC.ULTIMO.get('fechas'): dec = DC.trazabilidad(corte)
        elif os.path.exists(ruta): dec = pd.read_excel(ruta, sheet_name='Decretos')
    except Exception:
        dec = None
    info = pd.DataFrame([('Mes', f'{MESES[mes - 1]} {anio}'), ('Fecha de corte de decretos', f'{corte:%d-%m-%Y}'),
                         ('Actualizado', datetime.datetime.now().strftime('%d-%m-%Y %H:%M')), ('Rojos', r['rojos']), ('Rojos sin explicar', r['rojos_pendientes']),
                         ('Amarillos', r['amarillos']), ('Aceptadas con comentario', r['aceptadas']),
                         ('Decretos', 'Hoja Decretos: fecha aplicada a cada decreto y su origen (toma de razón DIPRES o, si no está identificado en DIPRES, la fecha de la planilla)'),
                         ('Nota', 'Archivo de revisión del agente. No es leído por Power BI y no modifica el consolidado.')], columns=['Dato', 'Valor'])
    try:
        with pd.ExcelWriter(ruta, engine='openpyxl') as w:
            info.to_excel(w, sheet_name='Resumen', index=False); df.to_excel(w, sheet_name='Alertas', index=False)
            ws = w.book['Alertas']
            anchos = {'A': 10, 'B': 34, 'C': 20, 'D': 34, 'E': 12, 'F': 50, 'G': 90, 'H': 14, 'I': 14, 'J': 14, 'K': 20, 'L': 50, 'M': 14, 'N': 16, 'O': 10}
            for c, a in anchos.items(): ws.column_dimensions[c].width = a
            for c in ws[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E5F')
            for row in ws.iter_rows(min_row=2):
                color = {'Rojo': 'F8D7DA', 'Amarillo': 'FFF3CD'}.get(row[0].value)
                if color: row[0].fill = PatternFill('solid', fgColor=color)
                for c in row[7:10]: c.number_format = '#,##0'
                row[6].alignment = Alignment(wrap_text=True, vertical='top')
            ws.column_dimensions['O'].hidden = True
            w.book['Resumen'].column_dimensions['A'].width = 30; w.book['Resumen'].column_dimensions['B'].width = 90
            if dec is not None and len(dec):
                dec.to_excel(w, sheet_name='Decretos', index=False); wd = w.book['Decretos']
                for c in wd[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E5F')
                for col, an in zip('ABCDEFGH', (11, 12, 14, 20, 20, 14, 58, 22)): wd.column_dimensions[col].width = an
                for row in wd.iter_rows(min_row=2):
                    for c in row[2:6]: c.number_format = 'DD-MM-YYYY'
                    if str(row[6].value or '').startswith('Planilla (no'): row[6].fill = PatternFill('solid', fgColor='EDEDED')
                wd.freeze_panes = 'A2'
    except PermissionError:
        raise PermissionError(f'No pude guardar {os.path.basename(ruta)}. ¿Está abierto en Excel? Ciérrelo e intente de nuevo.')
    return ruta

def usuario():
    try: return getpass.getuser()
    except Exception: return os.environ.get('USERNAME', '')

def aceptar(df, idx, comentario):
    df.loc[idx, ['Estado', 'Comentario', 'Usuario', 'Fecha']] = ['Aceptada', comentario.strip(), usuario(), datetime.datetime.now().strftime('%d-%m-%Y %H:%M')]

def quitar(df, idx):
    df.loc[idx, ['Estado', 'Comentario', 'Usuario', 'Fecha']] = ['Pendiente', '', '', '']


# ------------------------------------------------------------------ hoja de control
def hoja_de_control(df, ctrl, carpeta_mes, anio, mes, corte, consolidado, version):
    """Excel de una página para el visto bueno. Se niega si quedan rojos sin explicar."""
    r = resumen(df)
    if r['rojos_pendientes']:
        raise ValueError(f'Quedan {r["rojos_pendientes"]} alerta(s) en rojo sin explicar. Corríjalas o acéptelas con un comentario antes de exportar la hoja de control.')
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = Workbook(); ws = wb.active; ws.title = 'Hoja de control'
    azul = '1F4E5F'; neg = Font(bold=True); fino = Side(style='thin', color='BFBFBF')
    ws['A1'] = 'MINISTERIO DE AGRICULTURA – UPROGE'; ws['A1'].font = Font(bold=True, size=13, color=azul)
    ws['A2'] = f'Hoja de control · Informe de Ejecución Presupuestaria a {MESES[mes - 1]} {anio}'; ws['A2'].font = Font(bold=True, size=12)
    datos = [('Consolidado', os.path.basename(consolidado)), ('Fecha de corte de decretos', f'{corte:%d-%m-%Y}'),
             ('Generado por', f'{usuario()} · {datetime.datetime.now():%d-%m-%Y %H:%M} · agente versión {version}'),
             ('Controles', f'{r["rojos"]} rojos y {r["amarillos"]} amarillos; {r["aceptadas"]} aceptados con comentario; 0 rojos sin explicar')]
    f = 4
    for k, v in datos: ws.cell(f, 1, k).font = neg; ws.cell(f, 2, v); f += 1
    f += 1; ws.cell(f, 1, 'Ejecución por servicio (miles de $)').font = Font(bold=True, color=azul); f += 1
    hdr = ['Servicio', 'Presupuesto vigente', 'Devengo', '% avance']
    for j, h in enumerate(hdr, 1): c = ws.cell(f, j, h); c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor=azul)
    g = ctrl[ctrl.Tipo == 'Gastos']; f += 1
    for _, x in g.iterrows():
        ws.cell(f, 1, x.Servicio); ws.cell(f, 2, x['Presupuesto Vigente']).number_format = '#,##0'; ws.cell(f, 3, x['Monto Devengo']).number_format = '#,##0'
        ws.cell(f, 4, x['Monto Devengo'] / x['Presupuesto Vigente'] if x['Presupuesto Vigente'] else 0).number_format = '0.0%'; f += 1
    tv, td = g['Presupuesto Vigente'].sum(), g['Monto Devengo'].sum()
    for j, v in enumerate(['TOTAL MINAGRI', tv, td, td / tv if tv else 0], 1):
        c = ws.cell(f, j, v); c.font = neg
        if j in (2, 3): c.number_format = '#,##0'
        if j == 4: c.number_format = '0.0%'
    f += 2; ws.cell(f, 1, 'Alertas aceptadas con comentario').font = Font(bold=True, color=azul); f += 1
    acc = df[df.Estado.str.startswith('Aceptada')]
    for j, h in enumerate(['Nivel', 'Control', 'Servicio / programa', 'Comentario (usuario)'], 1):
        c = ws.cell(f, j, h); c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor=azul)
    f += 1
    if not len(acc): ws.cell(f, 1, 'Sin alertas aceptadas.'); f += 1
    for _, x in acc.iterrows():
        ws.cell(f, 1, x.Nivel); ws.cell(f, 2, f'{x.Control}: {x.Nombre}'.strip(': '))
        ws.cell(f, 3, f'{x.Servicio} {x.Programa}'.strip()); ws.cell(f, 4, f'{x.Comentario} ({x.Usuario}, {x.Estado.lower()})')
        for j in range(1, 5): ws.cell(f, j).alignment = Alignment(wrap_text=True, vertical='top')
        f += 1
    f += 1; ws.cell(f, 1, f'Amarillos sin comentario: {r["amarillos_pendientes"]} (detalle en "MINAGRI Revisión a {MESES[mes - 1]} {anio}.xlsx").'); f += 3
    ws.cell(f, 1, 'Revisado por (analista):'); ws.cell(f, 3, 'Visto bueno (jefatura UPROGE):'); f += 2
    ws.cell(f, 1, 'Nombre y fecha: ____________________'); ws.cell(f, 3, 'Nombre y fecha: ____________________')
    for col, w in zip('ABCD', (30, 48, 26, 70)): ws.column_dimensions[col].width = w
    ws.page_setup.orientation = 'landscape'; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True; ws.print_options.horizontalCentered = True
    ruta = ruta_hoja_control(carpeta_mes, anio, mes)
    try:
        wb.save(ruta)
    except PermissionError:
        raise PermissionError(f'No pude guardar {os.path.basename(ruta)}. ¿Está abierto en Excel? Ciérrelo e intente de nuevo.')
    return ruta
