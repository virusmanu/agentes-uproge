"""Reporte Mensualizado (versión 2.1): ejecución mes a mes del año, en un solo archivo HTML.

Arma la serie mensual con los consolidados de cada mes del año que ya están en las carpetas de informes (el mismo
archivo que lee Power BI), sin volver a leer los informes de los servicios:
  - mes del reporte: el consolidado que se acaba de generar (el mismo que usa "Generar Dashboard");
  - meses anteriores: el consolidado oficial publicado de cada mes y, si no existe, el del agente.
Devengo mensual = devengo acumulado del mes − devengo acumulado del mes anterior (en enero, el acumulado).
% Cumplimiento Total = devengo acumulado / presupuesto vigente del mes.
Nunca modifica los consolidados: solo los lee. Para no releer cada vez los meses cerrados, guarda un resumen de cada
consolidado en respaldos/reporte mensualizado (se rehace solo si el archivo cambió).
"""
import os, re, json, base64, pickle, datetime
import pandas as pd

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
PLANT = os.path.join(AQUI, 'plantillas')
CACHE = os.path.join(RAIZ, 'respaldos', 'reporte mensualizado')
MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
NOMBRE_SERV = {'SUBSE': 'SUBSECRETARÍA', 'ODEPA': 'ODEPA', 'INDAP': 'INDAP', 'SAG': 'SAG', 'CONAF': 'CONAF', 'CNR': 'CNR'}
MED = ['Presupuesto Ley', 'Presupuesto Vigente', 'Monto Devengo']
CLAVE = ['Servicio', 'Programa', 'Tipo', '_st', '_it', '_as']
EXCLUIR_SUBT = {35}   # saldo final de caja: fuera, igual que en el dashboard del informe


class ErrorMensualizado(Exception):
    pass


# ------------------------------------------------------------------ ubicar los consolidados de cada mes
def _carpeta_del_mes(carpeta_anio, anio, mes):
    if not os.path.isdir(carpeta_anio): return None
    for d in os.listdir(carpeta_anio):
        p = d.split()
        if len(p) >= 3 and p[0].isdigit() and int(p[0]) == mes and p[-1] == str(anio) and os.path.isdir(os.path.join(carpeta_anio, d)):
            return os.path.join(carpeta_anio, d)
    return None


def _buscar(carpeta, nombre):
    """Archivo de la carpeta con ese nombre, sin distinguir mayúsculas (los consolidados antiguos dicen "enero")."""
    if not carpeta or not os.path.isdir(carpeta): return None
    for f in os.listdir(carpeta):
        if f.lower() == nombre.lower(): return os.path.join(carpeta, f)
    return None


def consolidados(carpeta_mes, anio, mes, actual=None):
    """[(mes, ruta o None, fuente)] de enero al mes del reporte. fuente: 'oficial' | 'agente' | 'este corte' | 'falta'."""
    carpeta_anio = os.path.dirname(os.path.abspath(carpeta_mes))
    out = []
    for m in range(1, mes + 1):
        if m == mes and actual and os.path.exists(actual):
            out.append((m, actual, 'este corte')); continue
        cm = carpeta_mes if m == mes else _carpeta_del_mes(carpeta_anio, anio, m)
        d = os.path.join(cm, 'MINAGRI') if cm else None
        of = _buscar(d, f'MINAGRI Consolidado a {MESES[m - 1]} {anio}.xlsx'); ag = _buscar(d, f'MINAGRI Consolidado a {MESES[m - 1]} {anio} - AGENTE.xlsx')
        if of: out.append((m, of, 'oficial'))
        elif ag: out.append((m, ag, 'agente'))
        else: out.append((m, None, 'falta'))
    return out


# ------------------------------------------------------------------ resumen de un consolidado
def _num(s):
    return pd.to_numeric(s, errors='coerce').fillna(0).astype('int64')


def resumir(base):
    """Agrega la Base Consolidada a servicio, programa, tipo, subtítulo, ítem y asignación (sin región ni subasignación)."""
    b = base.copy()
    b['_st'] = _num(b['Subtítulo']); b['_it'] = _num(b['Ítem']); b['_as'] = _num(b['Asignación'])
    b.loc[b['_as'] % 1000 == 0, '_as'] = 0          # asignación "000" = nivel ítem (como en el informe)
    b = b[~b['_st'].isin(EXCLUIR_SUBT)]
    for m in MED: b[m] = pd.to_numeric(b[m], errors='coerce').fillna(0)
    b[['Servicio', 'Programa', 'Tipo']] = b[['Servicio', 'Programa', 'Tipo']].fillna('').astype(str)
    g = b.groupby(CLAVE)[MED].sum().reset_index()
    g = g[g[MED].abs().sum(axis=1) > 0.0005]
    nombres = {'st': {}, 'it': {}, 'as': {}}
    for c, n, k in (('_st', 'Nombre Subtítulo', 'st'), ('_it', 'Nombre Ítem', 'it'), ('_as', 'Nombre Asignación', 'as')):
        x = b[(b[c] > 0) & b[n].notna()][[c, n]].drop_duplicates(c)
        nombres[k] = {int(a): str(t) for a, t in x.values}
    return g, nombres


def _leer(ruta):
    """Resumen del consolidado, desde la copia guardada si el archivo no cambió."""
    st = os.stat(ruta); firma = (os.path.abspath(ruta), st.st_size, int(st.st_mtime))
    os.makedirs(CACHE, exist_ok=True)
    cp = os.path.join(CACHE, re.sub(r'[^\w.-]+', '_', os.path.basename(ruta)) + '.pkl')
    try:
        with open(cp, 'rb') as f:
            c = pickle.load(f)
        if c.get('firma') == firma and c.get('v') == 1: return c['g'], c['nombres']
    except Exception:
        pass
    try:
        base = pd.read_excel(ruta, sheet_name='Base Consolidada')
    except PermissionError:
        raise ErrorMensualizado(f'No se pudo leer {os.path.basename(ruta)}: está abierto o bloqueado. Ciérrelo e intente de nuevo.')
    except ValueError:
        raise ErrorMensualizado(f'{os.path.basename(ruta)} no tiene la hoja "Base Consolidada".')
    g, nombres = resumir(base)
    try:
        with open(cp, 'wb') as f:
            pickle.dump({'v': 1, 'firma': firma, 'g': g, 'nombres': nombres}, f)
    except Exception:
        pass
    return g, nombres


# ------------------------------------------------------------------ datos del reporte
def datos(cortes, anio, mes, version='', log=lambda s: None):
    """cortes: [(mes, ruta, fuente)]. Devuelve el diccionario que lee la plantilla."""
    n = len(cortes); por_mes = {}; nombres = {'st': {}, 'it': {}, 'as': {}}
    for m, ruta, fuente in cortes:
        if not ruta: continue
        log(f'  {MESES[m - 1]}: {fuente} ({os.path.basename(ruta)})')
        g, nm = _leer(ruta); por_mes[m] = g
        for k in nombres: nombres[k].update(nm[k])      # gana el nombre del mes más reciente
    if mes not in por_mes: raise ErrorMensualizado(f'No está el consolidado de {MESES[mes - 1]} {anio}.')
    todo = pd.concat([g.assign(_m=m) for m, g in por_mes.items()], ignore_index=True)
    serv = sorted(todo['Servicio'].unique(), key=lambda s: s[:2]); si = {s: i for i, s in enumerate(serv)}
    prog = sorted(todo['Programa'].unique()); pi = {p: i for i, p in enumerate(prog)}
    tipo = ['Gastos', 'Ingresos']; ti = {t: i for i, t in enumerate(tipo)}
    filas = []
    r3 = lambda v: round(float(v), 3)
    piv = todo.pivot_table(index=CLAVE, columns='_m', values=MED, aggfunc='sum', fill_value=0)
    for clave, r in piv.iterrows():
        s, p, t, st, it, as_ = clave
        if t not in ti: continue
        serie = lambda med: [r3(r[(med, m)]) if m in por_mes and (med, m) in r.index else None for m in range(1, n + 1)]
        filas.append([si[s], pi[p], ti[t], int(st), int(it), int(as_), serie('Presupuesto Ley'), serie('Presupuesto Vigente'), serie('Monto Devengo')])
    corto = lambda s: s.split(' ', 1)[1] if ' ' in s else s
    return {'anio': anio, 'mes': mes, 'meses': [MESES[m - 1] for m in range(1, n + 1)],
            'disponible': [m in por_mes for m in range(1, n + 1)],
            'fuentes': [{'mes': MESES[m - 1], 'fuente': f, 'archivo': os.path.basename(r) if r else ''} for m, r, f in cortes],
            'dic': {'serv': [NOMBRE_SERV.get(corto(s), corto(s)) for s in serv], 'prog': prog, 'tipo': tipo,
                    'st': {str(k): v for k, v in nombres['st'].items()}, 'it': {str(k): v for k, v in nombres['it'].items()},
                    'as': {str(k): v for k, v in nombres['as'].items()}},
            'filas': filas,
            'umbral': {'verde': mes / 12, 'amarillo': (mes - 1) / 12},
            'pie': f'Generado por el Agente de Ejecución Presupuestaria (versión {version}) el {datetime.datetime.now():%d-%m-%Y %H:%M}'}


def _b64(nombre):
    with open(os.path.join(PLANT, nombre), 'rb') as f:
        return 'data:image/png;base64,' + base64.b64encode(f.read()).decode('ascii')


def html(d, titulo):
    t = open(os.path.join(PLANT, 'mensualizado.html'), encoding='utf-8').read()
    js = json.dumps(d, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    t = t.replace('__TITULO__', titulo).replace('__LOGO_B__', _b64('logo_b.png'))
    return t.replace('__DATA__', js)


def ruta_reporte(carpeta_mes, anio, mes):
    return os.path.join(carpeta_mes, 'MINAGRI', f'{mes:02d} {MESES[mes - 1]} {anio} - Reporte Mensualizado MINAGRI.html')


def generar(consolidado, anio, mes, carpeta_mes, version='', log=lambda s: None):
    """Escribe el reporte del mes. Devuelve (ruta, cortes) para informar qué consolidado se usó en cada mes."""
    cortes = consolidados(carpeta_mes, anio, mes, actual=consolidado)
    d = datos(cortes, anio, mes, version, log)
    ruta = ruta_reporte(carpeta_mes, anio, mes)
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    with open(ruta, 'w', encoding='utf-8') as f:
        f.write(html(d, f'Reporte Mensualizado MINAGRI · {MESES[mes - 1]} {anio}'))
    return ruta, cortes
