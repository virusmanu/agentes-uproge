"""Fechas de los decretos de modificación presupuestaria según DIPRES.

Criterio UPROGE: un decreto se aplica al presupuesto vigente desde su TOMA DE RAZÓN por Contraloría.
Las planillas de Presupuesto Vigente de los servicios no siempre registran esa fecha (algunas anotan la fecha
del decreto), por lo que el agente la verifica en el PDF que publica DIPRES:
    https://www.dipres.gob.cl/Decretos/<año>/<N°>_TR.pdf
De ese PDF se leen la fecha del decreto ("SANTIAGO, 15 SEP 2026") y la del timbre "TOMADO RAZÓN" (con o sin
alcance). Si se configura una carpeta local de decretos, se busca primero ahí.

Lo encontrado se guarda en "Registro de decretos.xlsx" (junto al programa), para no volver a descargar.
En ese archivo se puede anotar a mano una fecha de toma de razón (columna "Fecha TR manual"), que prevalece.
Nada de esto modifica las planillas de Presupuesto Vigente ni el Catálogo Maestro: las fechas se aplican solo
en memoria al calcular el vigente a la fecha de corte."""
import os, re, sys, io, glob, time, datetime, unicodedata, urllib.request, urllib.error
import pandas as pd

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
sys.path.insert(0, os.path.join(AQUI, '_vendor'))          # lector de PDF incluido (pypdf, licencia BSD)
REGISTRO = os.path.join(RAIZ, 'Registro de decretos.xlsx')
URL = 'https://www.dipres.gob.cl/Decretos/{anio}/{n}_TR.pdf'
CRITERIOS = ('Toma de razón', 'Decreto', 'Planilla')
COLS = ['Año', 'N° Decreto', 'Fecha decreto', 'Fecha toma de razón', 'Con alcance', 'Fecha TR manual', 'Fuente', 'Consultado', 'Observación']

MESES = {'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6, 'jul': 7, 'ago': 8, 'sep': 9, 'set': 9, 'oct': 10, 'nov': 11, 'dic': 12}
ULTIMO = {}   # resultado de la última verificación: {'fechas': {N°: {...}}, 'criterio': str, 'avisos': [...]}


# ------------------------------------------------------------------ lectura del PDF
def _plano(t):
    t = unicodedata.normalize('NFKD', t)
    return ''.join(c for c in t if not unicodedata.combining(c)).upper()

_FECHA_TXT = r'(\d{1,2})\s*(?:DE\s+)?([A-Z]{3,10})\.?\s*(?:DE\s+|DEL\s+)?(\d{4})'
_FECHA_NUM = r'(\d{1,2})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{4})'

def _fecha(m_txt=None, m_num=None):
    try:
        if m_txt:
            d, mes, a = m_txt.groups(); mm = MESES.get(mes[:3].lower())
            return pd.Timestamp(int(a), mm, int(d)) if mm else None
        if m_num:
            d, mm, a = m_num.groups(); return pd.Timestamp(int(a), int(mm), int(d))
    except ValueError:
        return None

def _primera_fecha(t):
    """Primera fecha (texto o numérica) en el trozo t."""
    cand = []
    for m in re.finditer(_FECHA_TXT, t):
        f = _fecha(m_txt=m)
        if f is not None: cand.append((m.start(), f)); break
    for m in re.finditer(_FECHA_NUM, t):
        f = _fecha(m_num=m)
        if f is not None: cand.append((m.start(), f)); break
    return min(cand)[1] if cand else None

def leer_texto(texto):
    """Fechas del decreto y de la toma de razón desde el texto del PDF de DIPRES."""
    t = _plano(texto)
    r = {'fecha_decreto': None, 'fecha_tr': None, 'con_alcance': False}
    m = re.search(r'SANTIAGO\s*,?\s*' + _FECHA_TXT, t)
    if m: r['fecha_decreto'] = _fecha(m_txt=m)
    else:
        m = re.search(r'SANTIAGO\s*,?\s*' + _FECHA_NUM, t)
        if m: r['fecha_decreto'] = _fecha(m_num=m)
    m = re.search(r'T\s*O\s*M\s*A\s*D\s*[OA]\s*(?:DE\s*)?R\s*A\s*Z\s*O\s*N(\s*C\s*O\s*N\s*A\s*L\s*C\s*A\s*N\s*C\s*E)?', t)
    if m:
        r['con_alcance'] = bool(m.group(1))
        r['fecha_tr'] = _primera_fecha(t[m.end():m.end() + 250])
    return r

def leer_pdf(datos):
    """datos: bytes del PDF o ruta. Lee las primeras páginas y, si no aparece el timbre, el documento completo."""
    from pypdf import PdfReader
    rd = PdfReader(io.BytesIO(datos) if isinstance(datos, (bytes, bytearray)) else datos)
    pags = list(rd.pages); texto = ''
    for i, p in enumerate(pags):
        try: texto += (p.extract_text() or '') + '\n'
        except Exception: pass
        if i >= 1 and 'RAZON' in _plano(texto) and 'SANTIAGO' in _plano(texto): break
    r = leer_texto(texto); r['texto'] = texto
    return r


# ------------------------------------------------------------------ fuentes
def _en_carpeta(carpeta, n):
    if not carpeta or not os.path.isdir(carpeta): return []
    pat = re.compile(rf'(?<!\d)0*{n}(?!\d)')
    out = [f for f in glob.glob(os.path.join(carpeta, '**', '*.pdf'), recursive=True) if pat.search(os.path.basename(f))]
    return sorted(out, key=lambda f: ('TR' not in os.path.basename(f).upper(), len(f)))

def _descargar(anio, n, timeout):
    req = urllib.request.Request(URL.format(anio=anio, n=n), headers={'User-Agent': 'Agentes-UPROGE (MINAGRI)'})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()

def consultar(anio, n, carpeta_local='', en_linea=True, timeout=15):
    """Fechas de un decreto. Devuelve dict con fecha_decreto, fecha_tr, con_alcance, fuente, error.
    error='sin_conexion' si no se pudo llegar a DIPRES (para dejar de intentar con los demás)."""
    for f in _en_carpeta(carpeta_local, n):
        try:
            r = leer_pdf(f)
            if r['fecha_tr'] or r['fecha_decreto']:
                r.update(fuente='Carpeta local: ' + os.path.basename(f), error=''); return r
        except Exception:
            continue
    if not en_linea:
        return {'fecha_decreto': None, 'fecha_tr': None, 'con_alcance': False, 'fuente': '', 'error': 'sin_conexion'}
    try:
        r = leer_pdf(_descargar(anio, n, timeout)); r.update(fuente='DIPRES ' + URL.format(anio=anio, n=n), error='')
        if not r['fecha_tr']: r['error'] = 'sin_timbre'
        return r
    except urllib.error.HTTPError as e:
        return {'fecha_decreto': None, 'fecha_tr': None, 'con_alcance': False, 'fuente': '', 'error': f'no_publicado ({e.code})'}
    except (urllib.error.URLError, TimeoutError, OSError) as e:
        return {'fecha_decreto': None, 'fecha_tr': None, 'con_alcance': False, 'fuente': '', 'error': 'sin_conexion'}
    except Exception as e:
        return {'fecha_decreto': None, 'fecha_tr': None, 'con_alcance': False, 'fuente': '', 'error': f'pdf_ilegible ({type(e).__name__})'}


# ------------------------------------------------------------------ registro
def leer_registro():
    if not os.path.exists(REGISTRO): return pd.DataFrame(columns=COLS)
    r = pd.read_excel(REGISTRO, sheet_name='Decretos')
    for c in COLS:
        if c not in r: r[c] = None
    for c in ('Fecha decreto', 'Fecha toma de razón', 'Fecha TR manual'):
        r[c] = pd.to_datetime(r[c], errors='coerce', dayfirst=True)
    r['Año'] = pd.to_numeric(r['Año'], errors='coerce'); r['N° Decreto'] = pd.to_numeric(r['N° Decreto'], errors='coerce')
    return r[COLS]

def guardar_registro(r):
    from openpyxl.styles import Font, PatternFill
    r = r.sort_values(['Año', 'N° Decreto']).reset_index(drop=True)
    with pd.ExcelWriter(REGISTRO, engine='openpyxl') as w:
        r.to_excel(w, sheet_name='Decretos', index=False)
        ws = w.book['Decretos']
        for c in ws[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E5F')
        for col, ancho in zip('ABCDEFGHI', (7, 11, 14, 18, 11, 16, 60, 17, 60)): ws.column_dimensions[col].width = ancho
        for fila in ws.iter_rows(min_row=2):
            for c in (fila[2], fila[3], fila[5]): c.number_format = 'DD-MM-YYYY'
            fila[5].fill = PatternFill('solid', fgColor='FFF2CC')
        ws.freeze_panes = 'A2'
        ay = w.book.create_sheet('Cómo usar')
        for t in ('Registro de fechas de decretos de modificación presupuestaria (lo mantiene el agente).',
                  'El agente consulta cada decreto en DIPRES (o en la carpeta local configurada) y anota aquí la fecha del decreto y la de toma de razón.',
                  'Columna "Fecha TR manual" (amarilla): si DIPRES no tiene el PDF o la fecha leída es incorrecta, anote aquí la fecha de toma de razón; prevalece sobre la leída.',
                  'Para que el agente vuelva a consultar un decreto, borre su fila.',
                  'Este archivo no forma parte del informe de Power BI.'):
            ay.append([t])
        ay.column_dimensions['A'].width = 150

def decretos_en_planillas(pvs):
    """{N°: {servicio: fecha registrada en la planilla}} (Documento = Decreto)."""
    out = {}
    for s, ruta in pvs.items():
        try:
            b = pd.read_excel(ruta, sheet_name='Base Presupuesto', usecols=['Documento', 'Numero Documento', 'Fecha Documento'])
        except Exception:
            continue
        b = b[(b.Documento == 'Decreto') & b['Numero Documento'].notna()]
        b = b.assign(N=pd.to_numeric(b['Numero Documento'], errors='coerce'), F=pd.to_datetime(b['Fecha Documento'], errors='coerce'))
        for n, f in b[b.N.notna()].groupby('N')['F'].min().items():
            out.setdefault(int(n), {})[s] = f
    return out

def verificar(pvs, anio, criterio='Toma de razón', carpeta_local='', log=print, en_linea=True):
    """Completa el registro con los decretos de las planillas y devuelve las fechas a aplicar.
    Devuelve (aplicar, info, avisos): aplicar = {'1289': Timestamp, ...} según el criterio;
    info = {N°: {'planillas': {...}, 'decreto': ts, 'tr': ts, 'manual': ts, 'alcance': bool, 'fuente': str, 'error': str}}."""
    global ULTIMO
    criterio = criterio if criterio in CRITERIOS else 'Toma de razón'
    P = decretos_en_planillas(pvs); av = []
    reg = leer_registro(); nuevos = []; en_red = en_linea; faltan = []
    if criterio != 'Planilla':
        for n in sorted(P):
            f = reg[(reg['Año'] == anio) & (reg['N° Decreto'] == n)]
            if len(f) and (pd.notna(f['Fecha toma de razón'].iloc[0]) or pd.notna(f['Fecha TR manual'].iloc[0])):
                continue
            r = consultar(anio, n, carpeta_local, en_red)
            if en_red and r['fuente'].startswith('DIPRES'): time.sleep(0.4)       # pausa entre descargas
            if r['error'] == 'sin_conexion': en_red = False
            if r['error']: faltan.append((n, r['error']))
            fila = {'Año': anio, 'N° Decreto': n, 'Fecha decreto': r['fecha_decreto'], 'Fecha toma de razón': r['fecha_tr'],
                    'Con alcance': 'Sí' if r['con_alcance'] else ('No' if r['fecha_tr'] else ''), 'Fecha TR manual': pd.NaT,
                    'Fuente': r['fuente'], 'Consultado': datetime.datetime.now().strftime('%d-%m-%Y %H:%M'),
                    'Observación': '' if not r['error'] else {'sin_timbre': 'El PDF no tiene timbre de toma de razón legible'}.get(r['error'], r['error'])}
            if len(f):
                fila['Fecha TR manual'] = f['Fecha TR manual'].iloc[0]
                reg = reg.drop(f.index)
            nuevos.append(fila)
            log(f'Decreto N°{n}: ' + (f'toma de razón {r["fecha_tr"]:%d-%m-%Y}' if r['fecha_tr'] is not None else f'sin fecha ({r["error"]})'))
        if nuevos:
            reg = pd.concat([reg, pd.DataFrame(nuevos, columns=COLS)], ignore_index=True)
            try: guardar_registro(reg)
            except PermissionError: av.append(('Revisar', 'No pude guardar "Registro de decretos.xlsx" (¿está abierto en Excel?); las fechas se usaron igual en este proceso'))
    info, aplicar = {}, {}
    for n, plan in P.items():
        f = reg[(reg['Año'] == anio) & (reg['N° Decreto'] == n)]
        g = lambda c: (f[c].iloc[0] if len(f) and pd.notna(f[c].iloc[0]) else None)
        i = {'planillas': plan, 'decreto': g('Fecha decreto'), 'tr': g('Fecha toma de razón'), 'manual': g('Fecha TR manual'),
             'alcance': (g('Con alcance') == 'Sí'), 'fuente': g('Fuente') or '', 'error': g('Observación') or ''}
        tr = i['manual'] or i['tr']
        if criterio == 'Toma de razón' and tr is not None: aplicar[str(n)] = pd.Timestamp(tr)
        elif criterio == 'Decreto' and i['decreto'] is not None: aplicar[str(n)] = pd.Timestamp(i['decreto'])
        i['aplicada'] = aplicar.get(str(n)); info[n] = i
    if criterio != 'Planilla':
        sin = [n for n in sorted(P) if str(n) not in aplicar]
        if sin:
            # Normal: hay latencia entre la toma de razón y la publicación en DIPRES, y decretos recientes aún en trámite.
            # Esos decretos NO se descartan: se aplican con la fecha registrada en la planilla del servicio.
            av.append(('Info', ('Sin conexión con DIPRES: ' if (not en_red and en_linea) else '')
                       + f'{len(sin)} decreto(s) no identificados en DIPRES (en trámite de toma de razón o aún no publicados) se aplicaron con la fecha de la planilla: N°'
                       + ', N°'.join(map(str, sin)) + '. Detalle por servicio en la hoja Decretos del archivo de revisión del mes'))
        cambian = [(n, info[n]) for n in sorted(P) if str(n) in aplicar and any(pd.notna(fp) and fp != aplicar[str(n)] for fp in info[n]['planillas'].values())]
        if cambian:
            av.append(('Info', f'Fechas de decretos ajustadas al criterio "{criterio}" (fecha DIPRES ≠ fecha de la planilla): '
                       + '; '.join(f'N°{n} → {aplicar[str(n)]:%d-%m-%Y} (' + ', '.join(f'{s} {fp:%d-%m-%Y}' for s, fp in i['planillas'].items() if pd.notna(fp)) + ')'
                                   for n, i in cambian)))
    carpeta_info = {}
    if carpeta_local:
        try: carpeta_info = revisar_carpeta(carpeta_local, anio)
        except Exception as e: av.append(('Info', f'No se pudo revisar la carpeta de decretos ({type(e).__name__}: {e})'))
    ULTIMO = {'fechas': info, 'criterio': criterio, 'aplicar': aplicar, 'avisos': av, 'anio': anio, 'carpeta': carpeta_info,
              'en_planillas': {n: set(p) for n, p in P.items()}}
    return aplicar, info, av


def trazabilidad(corte=None):
    """Una fila por decreto y servicio: fecha de la planilla, fechas DIPRES, fecha aplicada y su origen."""
    info = (ULTIMO or {}).get('fechas', {}); crit = (ULTIMO or {}).get('criterio', '')
    c = pd.Timestamp(corte) if corte is not None else None; filas = []
    for n, i in sorted(info.items()):
        for s, fp in sorted(i['planillas'].items()):
            fp = pd.Timestamp(fp) if fp is not None and pd.notna(fp) else None
            ap = i.get('aplicada'); ap = pd.Timestamp(ap) if ap is not None else fp
            if crit == 'Planilla': origen = 'Planilla (criterio Planilla)'
            elif i.get('aplicada') is None: origen = 'Planilla (no identificado en DIPRES: en trámite o sin publicar)'
            elif i.get('manual') is not None and crit == 'Toma de razón': origen = 'Fecha TR manual (Registro de decretos)'
            else: origen = 'DIPRES, ' + ('toma de razón' if crit == 'Toma de razón' else 'fecha del decreto') + (' con alcance' if i.get('alcance') and crit == 'Toma de razón' else '')
            filas.append({'N° Decreto': n, 'Servicio': s, 'Fecha planilla': fp, 'Fecha decreto (DIPRES)': i.get('decreto'),
                          'Toma de razón (DIPRES)': i.get('manual') or i.get('tr'), 'Fecha aplicada': ap, 'Origen de la fecha aplicada': origen,
                          'Dentro del corte': ('' if c is None or ap is None else 'Sí' if ap <= c else 'No (posterior al corte)')})
    return pd.DataFrame(filas, columns=['N° Decreto', 'Servicio', 'Fecha planilla', 'Fecha decreto (DIPRES)', 'Toma de razón (DIPRES)',
                                        'Fecha aplicada', 'Origen de la fecha aplicada', 'Dentro del corte'])


# ------------------------------------------------------------------ decretos de la carpeta local
SERVICIOS_TEXTO = {'01 SUBSE': ('SUBSECRETARIA DE AGRICULTURA',), '02 ODEPA': ('OFICINA DE ESTUDIOS Y POLITICAS AGRARIAS',),
                   '03 INDAP': ('INSTITUTO DE DESARROLLO AGROPECUARIO',), '04 SAG': ('SERVICIO AGRICOLA Y GANADERO',),
                   '05 CONAF': ('CORPORACION NACIONAL FORESTAL',), '06 CNR': ('COMISION NACIONAL DE RIEGO',)}
CACHE = os.path.join(RAIZ, 'respaldos', 'decretos_leidos.json')

def servicios_del_texto(texto):
    """Servicios de la Partida 13 que nombra el decreto (por nombre o por capítulo 13 0X)."""
    t = re.sub(r'\s+', ' ', _plano(texto)); out = set()
    for s, nombres in SERVICIOS_TEXTO.items():
        if any(n in t for n in nombres) or re.search(rf'\b130{s[1]}\b', t): out.add(s)
    return out

def revisar_carpeta(carpeta, anio):
    """{N°: {'servicios', 'fecha_decreto', 'fecha_tr', 'archivo'}} de los PDF de la carpeta local de decretos del año.
    Lo leído se guarda en respaldos/decretos_leidos.json (se vuelve a leer solo si el PDF cambia)."""
    import json
    if not carpeta or not os.path.isdir(carpeta): return {}
    try: cache = json.load(open(CACHE, encoding='utf-8'))
    except Exception: cache = {}
    out, nuevo = {}, False
    for f in glob.glob(os.path.join(carpeta, '**', '*.pdf'), recursive=True):
        nums = [x for x in re.findall(r'(?<!\d)(\d{2,4})(?!\d)', os.path.basename(f)) if int(x) not in (anio - 1, anio, anio + 1)]
        if not nums: continue
        m = re.match(r'(\d+)', nums[0])
        clave = f'{f}|{os.path.getmtime(f)}'
        if clave not in cache:
            try:
                r = leer_pdf(f)
                cache[clave] = {'n': int(m.group(1)), 'servicios': sorted(servicios_del_texto(r['texto'])),
                                'decreto': r['fecha_decreto'].strftime('%Y-%m-%d') if r['fecha_decreto'] is not None else None,
                                'tr': r['fecha_tr'].strftime('%Y-%m-%d') if r['fecha_tr'] is not None else None}
            except Exception:
                cache[clave] = {'n': int(m.group(1)), 'servicios': [], 'decreto': None, 'tr': None}
            nuevo = True
        c = cache[clave]; fe = c['tr'] or c['decreto']
        if fe and int(fe[:4]) != anio: continue
        if not fe and str(anio) not in f: continue
        out[c['n']] = {'servicios': set(c['servicios']), 'fecha_decreto': pd.Timestamp(c['decreto']) if c['decreto'] else None,
                       'fecha_tr': pd.Timestamp(c['tr']) if c['tr'] else None, 'archivo': os.path.basename(f)}
    if nuevo:
        try:
            os.makedirs(os.path.dirname(CACHE), exist_ok=True); json.dump(cache, open(CACHE, 'w', encoding='utf-8'), ensure_ascii=False)
        except Exception: pass
    return out
