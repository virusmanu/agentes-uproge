"""Proceso completo del agente, común a la versión de consola (.bat) y a la versión gráfica (ventana).
Ninguna de las dos tiene lógica propia: ambas llaman a estas funciones, por lo que dan el mismo resultado."""
import os, configparser, datetime, warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
import pandas as pd
from motor import *
import ideas as IF
import revision as RV

def _version():
    try:
        import json
        with open(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'version.json'), encoding='utf-8') as f:
            return json.load(f)['version']
    except Exception:
        return '1.1'
VERSION = _version()
MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG = os.path.join(RAIZ, 'config.ini')
CATALOGO_POR_DEFECTO = os.path.join(RAIZ, 'Catálogo Maestro Ejecución MINAGRI 2026.xlsx')


# ------------------------------------------------------------------ configuración
def leer_config():
    c = configparser.ConfigParser()
    if not c.read(CONFIG, encoding='utf-8') or 'RUTAS' not in c:
        raise ErrorAgente(f'No encuentro config.ini en {RAIZ}')
    g = c['RUTAS']
    cat = g.get('catalogo', '').strip() or CATALOGO_POR_DEFECTO
    if not os.path.isabs(cat): cat = os.path.join(RAIZ, cat)
    return {'carpeta_informes': g.get('carpeta_informes', '').strip(),
            'carpeta_presupuesto_vigente': g.get('carpeta_presupuesto_vigente', '').strip(),
            'carpeta_informes_anio_anterior': g.get('carpeta_informes_anio_anterior', '').strip(),
            'catalogo': cat,
            'carpeta_actualizaciones': g.get('carpeta_actualizaciones', '').strip(),
            'carpeta_decretos': g.get('carpeta_decretos', '').strip(),
            'repositorio_github': g.get('repositorio_github', '').strip()}

def guardar_config(cfg):
    """Reescribe solo los valores de las rutas, conservando los comentarios de config.ini."""
    lineas = open(CONFIG, encoding='utf-8').read().splitlines() if os.path.exists(CONFIG) else ['[RUTAS]']
    vistos = set(); out = []
    for l in lineas:
        k = l.split('=', 1)[0].strip()
        if '=' in l and not l.lstrip().startswith(';') and k in cfg:
            v = '' if (k == 'catalogo' and os.path.normcase(cfg[k]) == os.path.normcase(CATALOGO_POR_DEFECTO)) else cfg[k]
            out.append(f'{k} = {v}'); vistos.add(k)
        else:
            out.append(l)
    for k, v in cfg.items():
        if k not in vistos: out.append(f'{k} = {v}')
    open(CONFIG, 'w', encoding='utf-8').write('\n'.join(out) + '\n')

def revisar_rutas(cfg):
    """Lista de problemas con las rutas (vacía si todo está bien)."""
    p = []
    if not os.path.isdir(cfg['carpeta_informes']): p.append('La carpeta de informes mensuales no existe o no está configurada.')
    if not os.path.isdir(cfg['carpeta_presupuesto_vigente']): p.append('La carpeta de Presupuesto Vigente no existe o no está configurada.')
    if not os.path.exists(cfg['catalogo']): p.append('No se encuentra el Catálogo Maestro.')
    return p


# ------------------------------------------------------------------ utilidades del mes
def carpeta_del_mes(base, anio, mes):
    if not os.path.isdir(base): raise ErrorAgente(f'La carpeta de informes no existe: {base}')
    c = [d for d in os.listdir(base) if d.startswith(f'{mes:02d} ') and str(anio) in d and os.path.isdir(os.path.join(base, d))]
    if len(c) != 1: raise ErrorAgente(f'No encuentro una única carpeta "{mes:02d} {MESES[mes-1]} {anio}" dentro de {base} (encontré: {c})')
    return os.path.join(base, c[0])

def carpeta_del_mes_o_none(base, anio, mes):
    try: return carpeta_del_mes(base, anio, mes)
    except ErrorAgente: return None

def meses_disponibles(base):
    """[(anio, mes, nombre carpeta)] de las carpetas 'MM Mes AAAA' existentes, más reciente primero."""
    out = []
    if os.path.isdir(base):
        for d in os.listdir(base):
            p = d.split()
            if len(p) >= 3 and p[0].isdigit() and p[-1].isdigit() and 1 <= int(p[0]) <= 12 and os.path.isdir(os.path.join(base, d)):
                out.append((int(p[-1]), int(p[0]), d))
    return sorted(out, reverse=True)

def oficial_del_mes(carpeta):
    d = os.path.join(carpeta, 'MINAGRI')
    if not os.path.isdir(d): return None
    o = [os.path.join(d, f) for f in os.listdir(d) if f.startswith('MINAGRI Consolidado a') and f.endswith('.xlsx') and 'AGENTE' not in f.upper() and not f.startswith('~$')]
    return o[0] if len(o) == 1 else None

def salida_del_mes(carpeta, anio, m):
    return os.path.join(carpeta, 'MINAGRI', f'MINAGRI Consolidado a {MESES[m-1]} {anio} - AGENTE.xlsx')

def leer_fecha(t):
    if isinstance(t, (pd.Timestamp, datetime.date)): return pd.Timestamp(t)
    for f in ('%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d'):
        try: return pd.Timestamp(datetime.datetime.strptime(str(t).strip(), f))
        except ValueError: pass
    raise ErrorAgente(f'Fecha no válida: "{t}". Use DD-MM-AAAA, por ejemplo 06-11-2026')

def estado_del_mes(cfg, anio, m, carpeta=None):
    """Revisión previa (rápida) para mostrar antes de generar: archivos, consolidado oficial y fecha de corte."""
    carpeta = carpeta or carpeta_del_mes(cfg['carpeta_informes'], anio, m)
    try:
        arch = ubicar_archivos(carpeta); falt = []
    except ErrorAgente as e:
        arch = {}; falt = [str(e).split('\n')[0]]
    K = Catalogo(cfg['catalogo'])
    mes = f'{anio}-{m:02d}'
    return {'carpeta': carpeta, 'archivos': arch, 'problemas_archivos': falt, 'oficial': oficial_del_mes(carpeta),
            'corte_catalogo': K.corte.get(mes), 'salida': salida_del_mes(carpeta, anio, m)}

def sugerir_corte_mes(cfg, oficial):
    return sugerir_corte(planillas_pv(cfg['carpeta_presupuesto_vigente']), oficial)

def guardar_corte_en_catalogo(catalogo, anio, m, fecha, obs):
    """Agrega (o actualiza) la fecha de corte del mes en la hoja 'Corte decretos' del Catálogo Maestro."""
    import openpyxl, copy
    wb = openpyxl.load_workbook(catalogo); ws = wb['Corte decretos']; mes = f'{anio}-{m:02d}'
    fila = next((i for i in range(2, ws.max_row + 1) if str(ws.cell(i, 1).value or '').strip()[:7] == mes), None)
    if fila is None:
        fila = ws.max_row + 1
        for j in range(1, 5): ws.cell(fila, j)._style = copy.copy(ws.cell(fila - 1, j)._style)
    ws.cell(fila, 1, mes); ws.cell(fila, 2, f'{MESES[m-1]} {anio}')
    ws.cell(fila, 3, fecha.to_pydatetime()); ws.cell(fila, 3).number_format = 'DD-MM-YYYY'; ws.cell(fila, 4, obs)
    try: wb.save(catalogo)
    except PermissionError: raise ErrorAgente('No pude guardar la fecha en el Catálogo Maestro. ¿Está abierto en Excel? Ciérrelo e intente de nuevo.')


# ------------------------------------------------------------------ proceso completo
def procesar(cfg, anio, m, corte, origen_corte, carpeta=None, log=print, paso=lambda n, total, texto: None):
    """Genera el consolidado del mes. Devuelve un diccionario con el resultado.
    log(texto): mensajes de avance. paso(n, total, texto): para barras de progreso."""
    nombre = f'{MESES[m-1]} {anio}'
    carpeta = carpeta or carpeta_del_mes(cfg['carpeta_informes'], anio, m)
    TOTAL = 24; n = [0]
    def avanzar(t):
        n[0] += 1; paso(min(n[0], TOTAL), TOTAL, t); log(t)
    avanzar('Leyendo el Catálogo Maestro y las planillas de Presupuesto Vigente')
    K = Catalogo(cfg['catalogo']); pvs = planillas_pv(cfg['carpeta_presupuesto_vigente'])
    avanzar('Verificando fechas de decretos en DIPRES (toma de razón)')
    import motor, decretos as DC
    crit = RV.criterio_fecha_decreto(); motor.FECHAS_DECRETO.clear(); av_dec = []
    try:
        aplicar, _, av_dec = DC.verificar(pvs, anio, crit, cfg.get('carpeta_decretos', ''), log=log)
        motor.FECHAS_DECRETO.update(aplicar)
    except Exception as e:
        av_dec = [('Revisar', f'No se pudieron verificar las fechas de los decretos en DIPRES ({type(e).__name__}: {e}); se usaron las fechas de las planillas')]
    dir_min = os.path.join(carpeta, 'MINAGRI'); os.makedirs(dir_min, exist_ok=True)
    oficial = oficial_del_mes(carpeta)

    def log_motor(t):
        t = t.strip()
        if t.startswith('Leyendo'): avanzar(t.replace(' ...', ''))
        elif t.startswith('Presupuesto vigente'): avanzar(t.replace(' ...', ''))
    base_df, bg, _, ctrl, av = generar(K, carpeta, pvs, corte, None, log=log_motor)
    av = list(av_dec) + list(av)

    avanzar('Redactando el borrador de Ideas Fuerza')
    prev = prev_g = None; fi_prev = 'sin mes anterior (enero)'; corte_prev = None; cprev = None
    if m > 1:
        cprev = carpeta_del_mes_o_none(cfg['carpeta_informes'], anio, m - 1)
        pprev = IF.consolidado_de(cprev) if cprev else None
        if pprev:
            prev, prev_g = IF.leer_consolidado(pprev); fi_prev = os.path.basename(pprev)
            mprev = f'{anio}-{m-1:02d}'
            if mprev in K.corte: corte_prev = K.corte[mprev]
            else:
                corte_prev = sugerir_corte(pvs, pprev)[0]
                av.append(('Info', f'Fecha de corte de {mprev} no está en el catálogo; para listar los decretos del mes se usó {corte_prev:%d-%m-%Y}, calculada del consolidado de ese mes'))
        else:
            fi_prev = 'no se encontró el consolidado del mes anterior'
            av.append(('Revisar', f'No encontré el consolidado del mes anterior en la carpeta MINAGRI de {MESES[m-2]}; las Ideas Fuerza no comparan con el mes anterior'))
    dec = IF.decretos_del_periodo(pvs, corte_prev, corte) if corte_prev is not None else {}
    anual, fuente_anual = IF.pct_anio_anterior(carpeta, cfg.get('carpeta_informes_anio_anterior', ''), anio, m)
    if not anual: av.append(('Info', f'Sin datos de {anio-1} para la comparación anual (configure la carpeta del año anterior o agregue el Cuadro Comparativo del mes)'))
    # filas sin Subtítulo reconocible: se informan y no entran al detalle por subtítulo/asignación
    sin_st = base_df[pd.to_numeric(base_df['Subtítulo'], errors='coerce').isna()]
    if len(sin_st):
        r = sin_st.groupby(['Servicio', 'Tipo'])[['Presupuesto Vigente', 'Monto Devengo']].sum()
        mm = lambda v: f'{v:,.0f}'.replace(',', '.')
        for (sv, tp), x in r.iterrows():
            nf = int(((sin_st.Servicio == sv) & (sin_st.Tipo == tp)).sum())
            av.append(('Revisar', f'{sv} {tp}: {nf} fila(s) sin Subtítulo (vigente {mm(x["Presupuesto Vigente"])} M$, '
                                   f'devengo {mm(x["Monto Devengo"])} M$). Revise la planilla de Presupuesto Vigente o el informe base: '
                                   'filas en blanco o con código no numérico'))
    try:
        ideas, respaldo = IF.generar_ideas(base_df, bg, prev, prev_g, anual, dec, m, anio)
    except Exception as e:
        import traceback
        log('No se pudo redactar el borrador de Ideas Fuerza; el consolidado se genera igual:\n' + traceback.format_exc())
        av.append(('Revisar', f'No se pudo redactar el borrador de Ideas Fuerza ({type(e).__name__}: {e}). La hoja Ideas Fuerza queda con la plantilla para completar a mano'))
        ideas = plantilla_ideas() if 'plantilla_ideas' in globals() else pd.DataFrame(columns=['Código Institución', 'Institución', 'Ideas Fuerza'])
        respaldo = pd.DataFrame()
    fi = f'BORRADOR generado por el agente. Mes anterior: {fi_prev}. Año anterior: {fuente_anual or "no disponible"}'
    ideas_of = pd.read_excel(oficial, sheet_name='Ideas Fuerza') if oficial else None

    val = det = None
    if oficial:
        avanzar('Comparando con el consolidado oficial del mes')
        val, det = validar(base_df, bg, oficial)
    avanzar('Guardando el archivo')
    salida = salida_del_mes(carpeta, anio, m)
    info = pd.DataFrame([('Mes', nombre), ('Generado', datetime.datetime.now().strftime('%d-%m-%Y %H:%M')), ('Versión del agente', VERSION),
                         ('Carpeta de informes', carpeta), ('Fecha de corte de decretos', f'{corte:%d-%m-%Y} ({origen_corte})'), ('Fecha de aplicación de decretos', crit + (' (verificada en DIPRES)' if crit != 'Planilla' else '')), ('Ideas Fuerza', fi),
                         ('Filas Base Consolidada', len(base_df)), ('Filas Base Glosas', len(bg))], columns=['Dato', 'Valor'])
    avdf = pd.DataFrame(av or [('OK', 'Sin observaciones')], columns=['Tipo', 'Aviso'])
    escribir_salida(salida, base_df, bg, ideas, respaldo, ideas_of, info, ctrl, avdf, val, det)
    avanzar('Revisando el consolidado (controles de devengo)')
    rev = None
    try:
        rev = RV.revisar(base_df, bg, prev, anio, m, carpeta, cprev, CUADRATURA, pvs, corte)
    except PermissionError as e:
        av.append(('Revisar', str(e)))
    except Exception as e:
        import traceback
        log('No se pudo completar la revisión; el consolidado quedó guardado igual:\n' + traceback.format_exc())
        av.append(('Revisar', f'No se pudo completar la revisión ({type(e).__name__}: {e}). El consolidado quedó guardado igual'))
    paso(TOTAL, TOTAL, 'Listo')
    return {'revision': rev, 'ruta_revision': RV.ruta_revision(carpeta, anio, m), 'anio': anio, 'mes': m,
            'salida': salida, 'carpeta': carpeta, 'ctrl': ctrl, 'avisos': av, 'validacion': val, 'detalle': det,
            'ideas': ideas, 'oficial': oficial, 'corte': corte, 'origen_corte': origen_corte, 'nombre': nombre}

def escribir_salida(salida, base_df, bg, ideas, respaldo, ideas_of, info, ctrl, avdf, val, det):
    from openpyxl.styles import Alignment
    try:
        with pd.ExcelWriter(salida, engine='openpyxl') as w:
            base_df.to_excel(w, sheet_name='Base Consolidada', index=False); bg.to_excel(w, sheet_name='Base Glosas', index=False)
            ideas.to_excel(w, sheet_name='Ideas Fuerza', index=False)
            respaldo.to_excel(w, sheet_name='Análisis mensual', index=False)
            if ideas_of is not None: ideas_of.to_excel(w, sheet_name='Ideas Fuerza oficial', index=False)
            info.to_excel(w, sheet_name='Control', index=False); ctrl.to_excel(w, sheet_name='Control', index=False, startrow=len(info) + 2)
            avdf.to_excel(w, sheet_name='Control', index=False, startrow=len(info) + len(ctrl) + 5)
            if val is not None:
                val.to_excel(w, sheet_name='Validación', index=False); det.to_excel(w, sheet_name='Validación', index=False, startrow=len(val) + 3)
            for ws in w.book.worksheets:
                for col in ws.columns:
                    ws.column_dimensions[col[0].column_letter].width = min(max(len(str(c.value or '')) for c in col[:200]) + 2, 90)
            for n in ('Ideas Fuerza', 'Ideas Fuerza oficial'):
                if n in w.book.sheetnames:
                    ws = w.book[n]; ws.column_dimensions['C'].width = 150
                    for c in ws['C']: c.alignment = Alignment(wrap_text=True, vertical='top')
            ws = w.book['Análisis mensual']
            for row in ws.iter_rows(min_row=2):
                for c in row:
                    h = ws.cell(1, c.column).value or ''
                    if '%' in h or h == 'Var. pp': c.number_format = '0.0%'
                    elif h in ('Vigente', 'Devengo', 'Vigente ant.', 'Devengo ant.', 'Devengo del mes'): c.number_format = '#,##0'
    except PermissionError:
        raise ErrorAgente(f'No pude guardar {os.path.basename(salida)}. ¿Está abierto en Excel? Ciérrelo y vuelva a generar.')
