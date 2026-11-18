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
def _catalogo_por_defecto():
    """El catálogo del año más reciente que esté junto al programa ("Catálogo Maestro Ejecución MINAGRI <año>.xlsx")."""
    import glob, re
    c = [f for f in glob.glob(os.path.join(RAIZ, 'Catálogo Maestro Ejecución MINAGRI *.xlsx')) if re.search(r'(\d{4})\.xlsx$', f)]
    return max(c, key=lambda f: re.search(r'(\d{4})\.xlsx$', f).group(1)) if c else os.path.join(RAIZ, 'Catálogo Maestro Ejecución MINAGRI 2026.xlsx')
CATALOGO_POR_DEFECTO = _catalogo_por_defecto()


# ------------------------------------------------------------------ configuración
def leer_config():
    c = configparser.ConfigParser()
    if not c.read(CONFIG, encoding='utf-8') or 'RUTAS' not in c:
        raise ErrorAgente(f'No encuentro config.ini en {RAIZ}')
    g = c['RUTAS']
    cat = g.get('catalogo', '').strip() or _catalogo_por_defecto()
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
            v = '' if (k == 'catalogo' and os.path.normcase(cfg[k]) == os.path.normcase(_catalogo_por_defecto())) else cfg[k]
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

def _abierto_en_excel(ruta):
    """Excel deja un archivo "~$nombre" junto al libro abierto."""
    d, b = os.path.split(ruta)
    return os.path.exists(os.path.join(d, '~$' + b[2:])) or os.path.exists(os.path.join(d, '~$' + b))

def informes_completos(carpeta):
    """True si en la carpeta del mes están los 12 informes (uno por servicio y tipo). Revisión rápida, sin abrir archivos."""
    import glob
    from motor import SERV
    if not os.path.isdir(carpeta): return False
    nombres = [os.path.basename(f).lower() for f in glob.glob(os.path.join(carpeta, '*.xls*')) if not os.path.basename(f).startswith('~$')]
    return all(sum(1 for b in nombres if b.startswith(sv[:2]) and sv.split()[1].lower() in b and t in b) == 1
               for sv in SERV for t in ('gastos', 'ingresos'))

def revision_previa(cfg, anio, m, carpeta=None):
    """Revisión de los insumos ANTES de generar. Devuelve [(nivel, texto)] con nivel 'Error' (impide generar) o 'Aviso'."""
    import glob, re
    from motor import SERV
    R = []
    carpeta = carpeta or carpeta_del_mes(cfg['carpeta_informes'], anio, m)
    if not os.path.isdir(carpeta):
        return [('Error', f'No existe la carpeta del mes: {carpeta}')]
    # 1. informes del mes: uno por servicio y tipo
    hall = {}
    for f in sorted(glob.glob(os.path.join(carpeta, '*.xls')) + glob.glob(os.path.join(carpeta, '*.xlsx'))):
        b = os.path.basename(f)
        if b.startswith('~$'): continue
        for sv in SERV:
            if b.startswith(sv[:2]) and sv.split()[1].lower() in b.lower():
                for t in ('Gastos', 'Ingresos'):
                    if t.lower() in b.lower(): hall.setdefault((sv, t), []).append(f)
    faltan = [f'{sv} {t}' for sv in SERV for t in ('Gastos', 'Ingresos') if (sv, t) not in hall]
    if faltan: R.append(('Error', f'Faltan {len(faltan)} informe(s): ' + ', '.join(faltan)))
    for (sv, t), v in hall.items():
        if len(v) > 1: R.append(('Error', f'{sv} {t}: hay más de un archivo (' + ' / '.join(os.path.basename(x) for x in v) + '); deje solo el del mes'))
    # 2. cada informe: mes del nombre, fecha del archivo, abierto en Excel, se puede leer
    inicio = datetime.datetime(anio, m, 1)
    for (sv, t), v in sorted(hall.items()):
        if len(v) != 1: continue
        f = v[0]; b = os.path.basename(f); bl = b.lower()
        otro = [x for i, x in enumerate(MESES) if i != m - 1 and re.search(rf'\b{x.lower()}\b', bl)]
        if otro and MESES[m - 1].lower() not in bl:
            R.append(('Error', f'{b}: el nombre dice {otro[0]} y la carpeta es de {MESES[m - 1]}. ¿Es el informe del mes correcto?'))
        if datetime.datetime.fromtimestamp(os.path.getmtime(f)) < inicio:
            R.append(('Aviso', f'{b}: el archivo es del {datetime.datetime.fromtimestamp(os.path.getmtime(f)):%d-%m-%Y}, anterior al mes informado. Revise que sea la descarga del cierre de {MESES[m - 1]}'))
        if _abierto_en_excel(f):
            R.append(('Aviso', f'{b} está abierto en Excel; si tiene cambios sin guardar, el agente lee la última versión guardada'))
        try:
            if f.lower().endswith('.xlsx'):
                import openpyxl; openpyxl.load_workbook(f, read_only=True).close()
            else:
                import xlrd; xlrd.open_workbook(f, on_demand=True).release_resources()
        except Exception as e:
            R.append(('Error', f'{b}: no se puede abrir ({type(e).__name__}). ¿Está dañado o protegido?'))
    # 3. planillas de Presupuesto Vigente
    try:
        pvs = planillas_pv(cfg['carpeta_presupuesto_vigente'])
    except Exception as e:
        R.append(('Error', str(e))); pvs = {}
    for sv, p in pvs.items():
        b = os.path.basename(p)
        try:
            x = base_presupuesto(p)
            n = int(x['_completado'].sum())
            if n: R.append(('Aviso', f'{b}: {n} fila(s) con fórmulas sin calcular (el agente toma el código desde el nombre). Para dejarla al día: ábrala, presione Ctrl+Alt+F9 y guarde'))
            sin_f = x[(x.Documento == 'Decreto') & x['Fecha Documento'].isna() & x['Monto M$'].notna()]
            if len(sin_f): R.append(('Aviso', f'{b}: {len(sin_f)} fila(s) de decreto sin fecha'))
        except Exception as e:
            R.append(('Error', f'{b}: no se puede leer la hoja "Base Presupuesto" ({type(e).__name__}: {e})'))
        if _abierto_en_excel(p):
            R.append(('Aviso', f'{b} está abierta en Excel; se usa la última versión guardada'))
    # 4. catálogo y archivo de salida
    if not os.path.exists(cfg.get('catalogo', '')):
        R.append(('Error', f'No se encuentra el Catálogo Maestro: {cfg.get("catalogo", "")}'))
    sal = salida_del_mes(carpeta, anio, m)
    if os.path.exists(sal) and _abierto_en_excel(sal):
        R.append(('Error', f'{os.path.basename(sal)} está abierto en Excel: ciérrelo para poder reemplazarlo'))
    return R

def estado_del_mes(cfg, anio, m, carpeta=None):
    """Revisión previa (rápida) para mostrar antes de generar: archivos, consolidado oficial y fecha de corte."""
    carpeta = carpeta or carpeta_del_mes(cfg['carpeta_informes'], anio, m)
    try:
        arch = ubicar_archivos(carpeta); falt = []
    except ErrorAgente as e:
        # se muestran igual los que sí están (los duplicados y faltantes los detalla la revisión previa)
        import glob
        from motor import SERV
        arch = {}
        for f in glob.glob(os.path.join(carpeta, '*.xls*')):
            b = os.path.basename(f).lower()
            if b.startswith('~$'): continue
            for sv in SERV:
                for t in ('Gastos', 'Ingresos'):
                    if b.startswith(sv[:2]) and sv.split()[1].lower() in b and t.lower() in b: arch.setdefault((sv, t), f)
        falt = [str(e).split('\n')[0]]
    K = Catalogo(cfg['catalogo'])
    mes = f'{anio}-{m:02d}'
    return {'carpeta': carpeta, 'archivos': arch, 'problemas_archivos': falt, 'oficial': oficial_del_mes(carpeta),
            'corte_catalogo': K.corte.get(mes), 'salida': salida_del_mes(carpeta, anio, m),
            'previa': revision_previa(cfg, anio, m, carpeta)}

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
    import seguridad as SG; SG.catalogo_modificado_por_el_agente(catalogo, f'fecha de corte {mes}')


# ------------------------------------------------------------------ proceso completo
def procesar(cfg, anio, m, corte, origen_corte, carpeta=None, log=print, paso=lambda n, total, texto: None):
    """Genera el consolidado del mes. Devuelve un diccionario con el resultado.
    log(texto): mensajes de avance. paso(n, total, texto): para barras de progreso."""
    nombre = f'{MESES[m-1]} {anio}'
    carpeta = carpeta or carpeta_del_mes(cfg['carpeta_informes'], anio, m)
    TOTAL = 26; n = [0]
    def avanzar(t):
        n[0] += 1; paso(min(n[0], TOTAL), TOTAL, t); log(t)
    avanzar('Leyendo el Catálogo Maestro y las planillas de Presupuesto Vigente')
    import seguridad as SG
    cambio_cat = None
    try: cambio_cat = SG.verificar_catalogo(cfg['catalogo'])        # catálogo protegido: cambios hechos fuera del agente
    except Exception: pass
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
    if cambio_cat:
        av.insert(0, ('Revisar', 'Catálogo Maestro: ' + cambio_cat))
        try:
            import bitacora as BT; BT.registrar_accion(cfg['carpeta_informes'], anio, m, 'Catálogo modificado fuera del agente', cambio_cat)
        except Exception: pass
    try: av = [('Revisar' if n == 'Error' else 'Info', 'Revisión previa: ' + t) for n, t in revision_previa(cfg, anio, m, carpeta)
               if 'abierto en Excel' not in t] + av
    except Exception: pass

    avanzar('Redactando el borrador de Ideas Fuerza')
    prev = prev_g = None; fi_prev = 'sin mes anterior (enero)'; corte_prev = None; cprev = None; pprev = None
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
    # validación del corte: contratos de entrada y salida, cuadratura punta a punta, motor sombra, coherencia y huella
    avanzar('Validando el corte (cuadraturas, motor sombra y estructura para Power BI)')
    import precision as PR
    V = None; extra = []; ruta_val = None; pbix_h = []
    salida = salida_del_mes(carpeta, anio, m)
    try:
        insumos = list(ubicar_archivos(carpeta).values()) + list(pvs.values()) + [cfg['catalogo']]
        V = PR.validar_corte(K, base_df, bg, ideas, pvs, corte, anio, m, insumos, VERSION, RV.leer_reglas()[1], pprev)
        for _, x in V['hallazgos'][V['hallazgos'].Criticidad == PR.ROJO].iterrows():
            det_x = x['Detalle'] + (f". Qué revisar: {x['Qué revisar']}" if x['Qué revisar'] else '')
            extra.append(RV._alerta(RV.ROJO, f"Validación: {x['Control']}", x['Servicio'] or 'MINAGRI', '', None, x['Grupo'], det_x))
        pbix_h = revisar_pbix_del_mes(carpeta, anio, m, base_df, bg)
    except Exception as e:
        import traceback
        log('No se pudo completar la validación del corte:\n' + traceback.format_exc())
        av.append(('Revisar', f'No se pudo completar la validación del corte ({type(e).__name__}: {e})'))
    if V is not None and cambio_cat:
        V['hallazgos'] = pd.concat([pd.DataFrame([PR._h('Contrato de entrada', 'Catálogo Maestro modificado fuera del agente', PR.AMARILLO, cambio_cat, '',
                                                         'Confirme que los cambios del catálogo fueron intencionales (respaldo en respaldos\\catálogo protegido)')]),
                                    V['hallazgos']], ignore_index=True)
    sem = PR.semaforo(V['hallazgos']) if V is not None else None
    avanzar('Guardando el archivo')
    info = pd.DataFrame([('Mes', nombre), ('Generado', datetime.datetime.now().strftime('%d-%m-%Y %H:%M')), ('Versión del agente', VERSION),
                         ('Carpeta de informes', carpeta), ('Fecha de corte de decretos', f'{corte:%d-%m-%Y} ({origen_corte})'), ('Fecha de aplicación de decretos', crit + (' (verificada en DIPRES)' if crit != 'Planilla' else '')), ('Ideas Fuerza', fi),
                         ('Filas Base Consolidada', len(base_df)), ('Filas Base Glosas', len(bg))]
                        + ([('Validación del corte', '; '.join(f'{r.Control}: {r.Estado}' for r in sem.itertuples())), ('Huella del consolidado', V['huella'])] if sem is not None else []),
                        columns=['Dato', 'Valor'])
    avdf = pd.DataFrame(av or [('OK', 'Sin observaciones')], columns=['Tipo', 'Aviso'])
    escribir_salida(salida, base_df, bg, ideas, respaldo, ideas_of, info, ctrl, avdf, val, det)
    if V is not None:
        try:
            ruta_val = PR.guardar(V, carpeta, anio, m, corte, VERSION, salida, pbix_h)
        except Exception as e:
            av.append(('Revisar', str(e)))
    avanzar('Revisando el consolidado (controles de devengo)')
    rev = None
    try:
        rev = RV.revisar(base_df, bg, prev, anio, m, carpeta, cprev, CUADRATURA, pvs, corte, extra)
    except PermissionError as e:
        av.append(('Revisar', str(e)))
    except Exception as e:
        import traceback
        log('No se pudo completar la revisión; el consolidado quedó guardado igual:\n' + traceback.format_exc())
        av.append(('Revisar', f'No se pudo completar la revisión ({type(e).__name__}: {e}). El consolidado quedó guardado igual'))
    res = {'base': base_df, 'prev': prev, 'anual': anual, 'dec': dec, 'version': VERSION, 'revision': rev, 'carpeta': carpeta,
           'anio': anio, 'mes': m, 'salida': salida, 'corte': corte, 'ideas': ideas}
    avanzar('Preparando Ideas Fuerza y resumen ejecutivo en Word')
    docs = {}
    try:
        import documentos as DOC
        docs = {'ideas': DOC.ideas_fuerza(res), 'resumen': DOC.resumen_ejecutivo(res)}
    except Exception as e:
        av.append(('Revisar', f'No se pudieron preparar los documentos Word ({type(e).__name__}: {e}); el consolidado quedó guardado igual'))
    try:   # base histórica local (ficha, comparador y decretos del Panel de análisis)
        import historia as HI
        HI.guardar_mes(anio, m, base_df, 'agente', salida, V.get('traza_pv') if V is not None else None)
    except Exception as e:
        av.append(('Info', f'No se pudo guardar el mes en la base histórica ({type(e).__name__}: {e})'))
    try:   # bitácora de la generación (quién, cuándo, con qué insumos)
        import bitacora as BT, decretos as DC
        ins = [('Informe del mes', p) for p in ubicar_archivos(carpeta).values()] + [('Presupuesto Vigente', p) for p in pvs.values()] \
              + [('Catálogo Maestro', cfg['catalogo']), ('Reglas de revisión', RV.REGLAS), ('Registro de decretos', DC.REGISTRO)]
        hd = SG.huella_datos(salida)          # para comprobar al publicar que las cifras no cambiaron
        BT.registrar_generacion(cfg['carpeta_informes'], anio, m, res, ins, crit, obs=f'Huella de datos: {hd}')
        rotas = BT.verificar_cadena(cfg['carpeta_informes'], anio)       # bitácora encadenada: filas alteradas a mano
        if rotas:
            av.append(('Revisar', f'Bitácora del agente {anio}: {len(rotas)} fila(s) alteradas o agregadas a mano (' + '; '.join(f'hoja {h}, fila {f}' for h, f, _ in rotas[:5])
                       + '). La bitácora debe modificarse solo desde el agente'))
    except Exception as e:
        av.append(('Info', f'No se pudo anotar la generación en la bitácora ({type(e).__name__}: {e})'))
    paso(TOTAL, TOTAL, 'Listo')
    return {'documentos': docs, 'base': base_df, 'prev': prev, 'anual': anual, 'dec': dec, 'version': VERSION, 'pendientes': [dict(x) for x in PENDIENTES], 'revision': rev, 'ruta_revision': RV.ruta_revision(carpeta, anio, m), 'anio': anio, 'mes': m,
            'salida': salida, 'carpeta': carpeta, 'ctrl': ctrl, 'avisos': av, 'validacion': val, 'detalle': det,
            'ideas': ideas, 'oficial': oficial, 'corte': corte, 'origen_corte': origen_corte, 'nombre': nombre,
            'validacion_corte': V['hallazgos'] if V is not None else None, 'semaforo': sem, 'ruta_validacion': ruta_val, 'pbix': pbix_h}

def pbix_del_mes(carpeta):
    """Informes Power BI (.pbix) de la carpeta del mes (o de su carpeta MINAGRI)."""
    import glob
    return sorted(f for d in (carpeta, os.path.join(carpeta, 'MINAGRI')) for f in glob.glob(os.path.join(d, '*.pbix')) if not os.path.basename(f).startswith('~$'))

def revisar_pbix_del_mes(carpeta, anio, m, base_df, bg, rutas=None):
    import pbix as PB
    out = []; rutas = rutas or pbix_del_mes(carpeta)
    for r in rutas:
        for h in PB.revisar(r, anio, m, base_df, bg):
            if len(rutas) > 1: h['Elemento'] = f'{os.path.basename(r)} · {h["Elemento"]}'
            out.append(h)
    return out

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
