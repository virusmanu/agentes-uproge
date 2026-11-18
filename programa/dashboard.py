"""Dashboard HTML autocontenido del Informe de Ejecución Presupuestaria (versión 2.0).

Reproduce las seis páginas del informe Power BI (Portada, Hitos Relevantes, Presupuesto Consolidado, Ejecución
Glosas, Programas y Transferencias y Presupuesto Regional) con las mismas posiciones, colores, filtros y medidas,
a partir del consolidado del mes. Es un solo archivo .html: se abre con doble clic en cualquier navegador y se puede
enviar por correo. El mapa usa OpenStreetMap y necesita internet; el resto funciona sin conexión.
No modifica el consolidado ni el .pbix.
"""
import os, json, base64, datetime
import pandas as pd

AQUI = os.path.dirname(os.path.abspath(__file__))
PLANT = os.path.join(AQUI, 'plantillas')
MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

# capitales regionales (lat, lon) para el mapa de burbujas, por el valor de la columna Mapa del consolidado
COORDS = {
    'Región de Arica y Parinacota': (-18.48, -70.31), 'Región de Tarapacá': (-20.21, -70.15), 'Región de Antofagasta': (-23.65, -70.40),
    'Región de Atacama': (-27.37, -70.33), 'Región de Coquimbo': (-29.95, -71.34), 'Región de Valparaíso': (-33.05, -71.62),
    'Región Metropolitana de Santiago': (-33.45, -70.66), "Región del Libertador General Bernardo O'Higgins": (-34.17, -70.74),
    'Región del Maule': (-35.43, -71.66), 'Región de Ñuble': (-36.61, -72.10), 'Región del Biobío': (-36.83, -73.05),
    'Región de La Araucanía': (-38.74, -72.60), 'Región de Los Ríos': (-39.81, -73.25), 'Región de Los Lagos': (-41.47, -72.94),
    'Región de Aysén del General Carlos Ibáñez del Campo': (-45.57, -72.07), 'Región de Magallanes y de la Antártica Chilena': (-53.16, -70.91),
}


def _b64(nombre, mime='image/png'):
    with open(os.path.join(PLANT, nombre), 'rb') as f:
        return f'data:{mime};base64,' + base64.b64encode(f.read()).decode('ascii')


def _dic(serie, vacio=True):
    vals = sorted({str(v) for v in serie.dropna() if str(v).strip()})
    lista = ([''] if vacio else []) + vals
    idx = {v: i for i, v in enumerate(lista)}
    return lista, idx


def datos(base, bg, ideas, anio, mes, version=''):
    """Prepara los datos del dashboard (agregados y codificados) a partir de las hojas del consolidado."""
    b = base.copy()
    b['_serv'] = b['Servicio'].astype(str).str.split(' ', n=1).str[1]
    b['_cod'] = b['Servicio'].astype(str).str[:2]
    serv = [s for _, s in sorted(set(zip(b['_cod'], b['_serv'])))]
    si = {s: i for i, s in enumerate(serv)}
    prog, pi = _dic(b['Programa'], vacio=False)
    tipo = ['Gastos', 'Ingresos']; ti = {t: i for i, t in enumerate(tipo)}
    st, sti = _dic(b['Nombre Subtítulo']); it, iti = _dic(b['Nombre Ítem']); asg, ai = _dic(b['Nombre Asignación'])
    mp, mi = _dic(b['Mapa'])
    regiones = {int(c): r for c, r in b[['Código', 'Región']].dropna().drop_duplicates().values}
    for c in ('Presupuesto Ley', 'Presupuesto Vigente', 'Monto Requerimiento', 'Monto Devengo'): b[c] = pd.to_numeric(b[c], errors='coerce').fillna(0)
    k = ['_serv', 'Programa', 'Tipo', 'Nombre Subtítulo', 'Nombre Ítem', 'Nombre Asignación', 'Código', 'Mapa']
    g = b.fillna({c: '' for c in k if c != 'Código'}).fillna({'Código': 0}).groupby(k, dropna=False)[['Presupuesto Ley', 'Presupuesto Vigente', 'Monto Requerimiento', 'Monto Devengo']].sum().reset_index()
    g = g[g[['Presupuesto Ley', 'Presupuesto Vigente', 'Monto Requerimiento', 'Monto Devengo']].abs().sum(axis=1) > 0]
    r3 = lambda v: round(float(v), 3)
    filas = [[si[r['_serv']], pi[r['Programa']], ti.get(r['Tipo'], 0), sti.get(str(r['Nombre Subtítulo']), 0), iti.get(str(r['Nombre Ítem']), 0),
              ai.get(str(r['Nombre Asignación']), 0), int(r['Código'] or 0), mi.get(str(r['Mapa']), 0),
              r3(r['Presupuesto Ley']), r3(r['Presupuesto Vigente']), r3(r['Monto Requerimiento']), r3(r['Monto Devengo'])] for _, r in g.iterrows()]
    # glosas
    gl = bg.copy(); gl['_serv'] = gl['Servicio'].astype(str).str.split(' ', n=1).str[1]
    for s in gl['_serv'].unique():
        if s not in si: si[s] = len(serv); serv.append(s)
    for p in gl['Programa'].dropna().unique():
        if p not in pi: pi[p] = len(prog); prog.append(p)
    glosas = sorted(gl['Glosa'].dropna().astype(str).unique()); gi = {x: i for i, x in enumerate(glosas)}
    cod = gl.dropna(subset=['Glosa']).groupby('Glosa')['Código Glosa'].min()
    for c in ('Ley de Presupuestos', 'Presupuesto Vigente', 'Monto Devengo'): gl[c] = pd.to_numeric(gl[c], errors='coerce').fillna(0)
    gg = gl.dropna(subset=['Glosa']).groupby(['_serv', 'Programa', 'Glosa'])[['Ley de Presupuestos', 'Presupuesto Vigente', 'Monto Devengo']].sum().reset_index()
    gfil = [[si[r['_serv']], pi[r['Programa']], gi[str(r['Glosa'])], r3(r['Ley de Presupuestos']), r3(r['Presupuesto Vigente']), r3(r['Monto Devengo'])] for _, r in gg.iterrows()]
    ide = []
    if ideas is not None and len(ideas):
        for _, r in ideas.iterrows():
            try: ide.append({'c': int(r['Código Institución']), 'i': str(r['Institución']), 't': '' if pd.isna(r['Ideas Fuerza']) else str(r['Ideas Fuerza'])})
            except Exception: pass
    return {'anio': anio, 'mes': mes, 'mesNombre': MESES[mes - 1], 'mesMin': MESES[mes - 1].lower(),
            'umbral': {'verde': round(mes / 12, 3), 'amarillo': round((mes - 1) / 12, 3)},
            'dic': {'serv': serv, 'prog': prog, 'tipo': tipo, 'st': st, 'it': it, 'as': asg, 'mapa': mp, 'reg': {str(c): r for c, r in regiones.items()}, 'glosa': glosas},
            'codGlosa': [float(cod.get(x, 999)) if pd.notna(cod.get(x, 999)) else 999 for x in glosas],
            'filas': filas, 'glosas': gfil, 'ideas': ide, 'coords': {k: list(v) for k, v in COORDS.items()},
            'pie': f'Generado por el Agente de Ejecución Presupuestaria (versión {version}) el {datetime.datetime.now():%d-%m-%Y %H:%M}'}


def html(d, titulo):
    t = open(os.path.join(PLANT, 'dashboard.html'), encoding='utf-8').read()
    lf = os.path.join(AQUI, '_vendor', 'leaflet')
    js = json.dumps(d, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    rep = {'__TITULO__': titulo, '__LEAFLET_CSS__': open(os.path.join(lf, 'leaflet.css'), encoding='utf-8').read(),
           '__LEAFLET_JS__': open(os.path.join(lf, 'leaflet.js'), encoding='utf-8').read().replace('</script', '<\\/script'),
           '__LOGO_A__': _b64('logo_a.png'), '__LOGO_B__': _b64('logo_b.png'), '__FRANJA__': _b64('franja.png')}
    for k, v in rep.items(): t = t.replace(k, v)
    return t.replace('__DATA__', js)


def ruta_dashboard(carpeta_mes, anio, mes):
    return os.path.join(carpeta_mes, 'MINAGRI', f'{mes:02d} {MESES[mes - 1]} {anio} - Informe de Ejecución Presupuestaria MINAGRI.html')


def generar(consolidado, anio, mes, carpeta_mes, version=''):
    """Lee el consolidado del mes (hojas Base Consolidada, Base Glosas e Ideas Fuerza) y escribe el dashboard."""
    base = pd.read_excel(consolidado, sheet_name='Base Consolidada')
    bg = pd.read_excel(consolidado, sheet_name='Base Glosas')
    try: ideas = pd.read_excel(consolidado, sheet_name='Ideas Fuerza')
    except Exception: ideas = None
    d = datos(base, bg, ideas, anio, mes, version)
    ruta = ruta_dashboard(carpeta_mes, anio, mes)
    os.makedirs(os.path.dirname(ruta), exist_ok=True)
    with open(ruta, 'w', encoding='utf-8') as f:
        f.write(html(d, f'Informe de Ejecución Presupuestaria MINAGRI · {MESES[mes - 1]} {anio}'))
    return ruta
