"""Base histórica local del agente (versión 2.0): todos los cortes mensuales en una base SQLite.

Cada generación guarda su consolidado (agregado por servicio, programa, tipo, clasificador y región) y los decretos del
mes con su fecha aplicada. La base se puede reconstruir en cualquier momento desde los consolidados guardados en las
carpetas de los meses (botón "Reconstruir"), por lo que nunca es la única copia de nada: es un índice para consultar
la historia sin abrir archivos (Panel de análisis: ficha por asignación, comparador entre meses y decretos).
Vive en la carpeta respaldos del agente (no en SharePoint, donde una base SQLite se puede dañar al sincronizar).
"""
import os, sqlite3, datetime, glob
import pandas as pd

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BD = os.path.join(RAIZ, 'respaldos', 'Base histórica.sqlite')
MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
MED = ['Presupuesto Ley', 'Presupuesto Vigente', 'Monto Requerimiento', 'Monto Devengo']
CLAVE = ['Servicio', 'Programa', 'Tipo', 'Nombre Subtítulo', 'Nombre Ítem', 'Nombre Asignación', 'Código']


def conectar():
    os.makedirs(os.path.dirname(BD), exist_ok=True)
    c = sqlite3.connect(BD)
    c.executescript("""
    CREATE TABLE IF NOT EXISTS cortes (anio INT, mes INT, fuente TEXT, archivo TEXT, cargado TEXT, PRIMARY KEY (anio, mes));
    CREATE TABLE IF NOT EXISTS ejecucion (anio INT, mes INT, servicio TEXT, programa TEXT, tipo TEXT, subtitulo TEXT, item TEXT, asignacion TEXT,
                                          region INT, ley REAL, vigente REAL, requerimiento REAL, devengo REAL);
    CREATE INDEX IF NOT EXISTS ix_ej ON ejecucion (anio, mes);
    CREATE TABLE IF NOT EXISTS decretos (anio INT, mes INT, servicio TEXT, numero TEXT, fecha_planilla TEXT, fecha_aplicada TEXT, incluido TEXT,
                                         gastos REAL, ingresos REAL);
    """)
    return c


def _agregar(base):
    b = base.copy()
    for m in MED: b[m] = pd.to_numeric(b[m], errors='coerce').fillna(0)
    g = b.fillna({c: '' for c in CLAVE if c != 'Código'}).fillna({'Código': 0}).groupby(CLAVE, dropna=False)[MED].sum().reset_index()
    return g[g[MED].abs().sum(axis=1) > 0]


def guardar_mes(anio, mes, base, fuente, archivo, traza_pv=None):
    g = _agregar(base)
    with conectar() as c:
        c.execute('DELETE FROM ejecucion WHERE anio=? AND mes=?', (anio, mes))
        c.executemany('INSERT INTO ejecucion VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)',
                      [(anio, mes, r[0], r[1], r[2], r[3], r[4], r[5], int(r[6] or 0), float(r[7]), float(r[8]), float(r[9]), float(r[10])) for r in g[CLAVE + MED].itertuples(index=False)])
        c.execute('INSERT OR REPLACE INTO cortes VALUES (?,?,?,?,?)', (anio, mes, fuente, os.path.basename(archivo), datetime.datetime.now().strftime('%d-%m-%Y %H:%M')))
        if traza_pv is not None and len(traza_pv):
            d = traza_pv[(traza_pv.Documento == 'Decreto') & traza_pv['Numero Documento'].notna()].copy()
            d['Monto M$'] = pd.to_numeric(d['Monto M$'], errors='coerce').fillna(0)
            f = lambda x: '' if pd.isna(x) else pd.Timestamp(x).strftime('%d-%m-%Y')
            filas = []
            for (s, n), x in d.groupby(['Servicio', 'Numero Documento']):
                filas.append((anio, mes, s, str(int(n)), f(x['Fecha Planilla'].min()) if 'Fecha Planilla' in x else f(x['Fecha Documento'].min()), f(x['Fecha Documento'].min()),
                              x['Incluido en el vigente'].iloc[0], float(x.loc[x.Tipo == 'Gastos', 'Monto M$'].sum()), float(x.loc[x.Tipo == 'Ingresos', 'Monto M$'].sum())))
            c.execute('DELETE FROM decretos WHERE anio=? AND mes=?', (anio, mes))
            c.executemany('INSERT INTO decretos VALUES (?,?,?,?,?,?,?,?,?)', filas)


def consolidado_del_mes(carpeta_mes, anio, mes):
    """El consolidado a guardar: el oficial publicado si existe; si no, el del agente. (ruta, fuente) o (None, None)."""
    d = os.path.join(carpeta_mes, 'MINAGRI')
    of = os.path.join(d, f'MINAGRI Consolidado a {MESES[mes - 1]} {anio}.xlsx'); ag = os.path.join(d, f'MINAGRI Consolidado a {MESES[mes - 1]} {anio} - AGENTE.xlsx')
    if os.path.exists(of): return of, 'oficial'
    if os.path.exists(ag): return ag, 'agente'
    return None, None


def reconstruir(carpetas, log=print):
    """Vuelve a cargar todos los meses que encuentre en las carpetas de informes (año actual y anterior)."""
    n = 0
    for base in [c for c in carpetas if c and os.path.isdir(c)]:
        for d in sorted(os.listdir(base)):
            p = d.split()
            if len(p) < 3 or not p[0].isdigit() or not p[-1].isdigit() or not os.path.isdir(os.path.join(base, d)): continue
            mes, anio = int(p[0]), int(p[-1])
            if not 1 <= mes <= 12: continue
            ruta, fuente = consolidado_del_mes(os.path.join(base, d), anio, mes)
            if not ruta: continue
            try:
                log(f'  {MESES[mes - 1]} {anio} ({fuente})')
                guardar_mes(anio, mes, pd.read_excel(ruta, sheet_name='Base Consolidada'), fuente, ruta); n += 1
            except Exception as e:
                log(f'  {MESES[mes - 1]} {anio}: no se pudo leer ({type(e).__name__})')
    return n


def meses():
    with conectar() as c:
        return c.execute('SELECT anio, mes, fuente, archivo, cargado FROM cortes ORDER BY anio, mes').fetchall()


def tabla(sql, *args):
    with conectar() as c:
        return pd.read_sql_query(sql, c, params=args)


# ------------------------------------------------------------------ panel de análisis (HTML)
def panel(destino, version=''):
    """Escribe el Panel de análisis (ficha por asignación, comparador entre meses y decretos) desde la base histórica."""
    import json, base64
    ms = meses()
    idx = {(a, m): i for i, (a, m, *_) in enumerate(ms)}
    e = tabla('SELECT anio, mes, servicio, programa, tipo, subtitulo, item, asignacion, SUM(vigente) v, SUM(devengo) d FROM ejecucion '
              'GROUP BY anio, mes, servicio, programa, tipo, subtitulo, item, asignacion')
    claves = {}; series = []
    for r in e.itertuples(index=False):
        k = (r.servicio, r.programa, r.tipo, r.subtitulo, r.item, r.asignacion)
        if k not in claves: claves[k] = len(claves); series.append([None] * len(ms))
        series[claves[k]][idx[(r.anio, r.mes)]] = [round(r.v, 3), round(r.d, 3)]
    dec = tabla('SELECT anio, mes, servicio, numero, fecha_planilla, fecha_aplicada, incluido, gastos, ingresos FROM decretos ORDER BY anio, mes, CAST(numero AS INT), servicio')
    d = {'meses': [[a, m, f] for a, m, f, *_ in ms], 'claves': [list(k) for k in claves], 'series': series, 'decretos': dec.values.tolist()}
    plant = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'plantillas')
    logo = 'data:image/png;base64,' + base64.b64encode(open(os.path.join(plant, 'logo_b.png'), 'rb').read()).decode()
    t = open(os.path.join(plant, 'panel.html'), encoding='utf-8').read()
    t = t.replace('__LOGO_B__', logo).replace('__PIE__', f'{len(ms)} mes(es) · actualizado el {datetime.datetime.now():%d-%m-%Y %H:%M} · versión {version}')
    t = t.replace('__DATA__', json.dumps(d, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/'))
    with open(destino, 'w', encoding='utf-8') as f: f.write(t)
    return destino
