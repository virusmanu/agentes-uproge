"""Revisión del informe Power BI (.pbix) del mes, solo lectura.

El agente no modifica el .pbix. Revisa lo que cambia de un mes a otro y suele quedar del mes anterior al copiar
el archivo, y que los filtros del informe encuentren sus valores en el consolidado:
  - Títulos y textos con el mes y año del informe.
  - Umbrales del semáforo de avance (formato condicional): mes/12 para el verde y (mes-1)/12 para el amarillo.
  - Valores fijos de los filtros (p. ej. '24 TRANSFERENCIAS CORRIENTES' o 'Viático Nacional') presentes en los datos.
El origen de datos (ruta del Excel) queda dentro del modelo comprimido y no se puede leer: se recuerda revisarlo.
"""
import json, re, zipfile

ROJO, AMARILLO, VERDE = 'Rojo', 'Amarillo', 'Verde'
MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
RX_MES = re.compile(r'\b(' + '|'.join(MESES) + r')\b(?:\s+de)?\s+(\d{4})', re.I)


def _h(crit, elemento, detalle):
    return {'Criticidad': crit, 'Elemento': elemento, 'Detalle': detalle}


def _visuales(z):
    """[(pagina, id_visual, objeto JSON)] para el formato nuevo (Report/definition) y el antiguo (Report/Layout)."""
    nombres = z.namelist(); out = []
    paginas = {}
    for n in nombres:
        m = re.match(r'Report/definition/pages/([^/]+)/page\.json$', n)
        if m:
            try: paginas[m.group(1)] = json.loads(z.read(n).decode('utf-8-sig')).get('displayName', m.group(1))
            except Exception: pass
    for n in nombres:
        m = re.match(r'Report/definition/pages/([^/]+)/(?:visuals/([^/]+)/visual|page)\.json$', n)
        if m:
            try: out.append((paginas.get(m.group(1), m.group(1)), m.group(2) or 'página', json.loads(z.read(n).decode('utf-8-sig'))))
            except Exception: pass
    if not out and 'Report/Layout' in nombres:
        L = json.loads(z.read('Report/Layout').decode('utf-16-le'))
        for s in L.get('sections', []):
            pag = s.get('displayName', '')
            try: out.append((pag, 'página', json.loads(s.get('filters') or '[]')))
            except Exception: pass
            for v in s.get('visualContainers', []):
                obj = {}
                for k in ('config', 'filters', 'query'):
                    try: obj[k] = json.loads(v.get(k) or '{}')
                    except Exception: pass
                out.append((pag, obj.get('config', {}).get('name', ''), obj))
    return out


def _strings(o, acc):
    if isinstance(o, dict):
        for v in o.values(): _strings(v, acc)
    elif isinstance(o, list):
        for v in o: _strings(v, acc)
    elif isinstance(o, str): acc.append(o)


def _filtros_in(o, acc, entidades):
    """Condiciones "In" (no las "Not In"): (entidad, columna, [valores])."""
    if isinstance(o, dict):
        if 'From' in o and isinstance(o['From'], list):
            for f in o['From']:
                if isinstance(f, dict) and 'Name' in f and 'Entity' in f: entidades[f['Name']] = f['Entity']
        if 'Not' in o: return
        if 'In' in o and isinstance(o['In'], dict):
            ex = o['In'].get('Expressions', []); vals = o['In'].get('Values', [])
            if len(ex) == 1 and 'Column' in ex[0]:
                col = ex[0]['Column']; prop = col.get('Property'); src = col.get('Expression', {}).get('SourceRef', {})
                ent = src.get('Entity') or entidades.get(src.get('Source'))
                lits = []
                for v in vals:
                    if v and isinstance(v[0], dict) and 'Literal' in v[0]:
                        lv = v[0]['Literal'].get('Value', '')
                        if lv.startswith("'") and lv.endswith("'"): lits.append(lv[1:-1].replace("''", "'"))
                if lits: acc.append((ent, prop, lits))
            return
        for v in o.values(): _filtros_in(v, acc, entidades)
    elif isinstance(o, list):
        for v in o: _filtros_in(v, acc, entidades)


def _umbrales(o, acc):
    """Literales decimales comparados con la medida 'Porcentaje de Avance' (formato condicional)."""
    if isinstance(o, dict):
        c = o.get('Comparison')
        if isinstance(c, dict):
            izq = c.get('Left', {}); der = c.get('Right', {})
            med = izq.get('Measure', {}).get('Property') if isinstance(izq, dict) else None
            lit = der.get('Literal', {}).get('Value') if isinstance(der, dict) else None
            if med and 'Avance' in med and isinstance(lit, str) and lit.endswith('D'):
                try:
                    x = float(lit[:-1])
                    if 0 < x < 1: acc.append((med, x))
                except ValueError: pass
        for v in o.values(): _umbrales(v, acc)
    elif isinstance(o, list):
        for v in o: _umbrales(v, acc)


def revisar(ruta, anio, mes, base=None, bg=None):
    H = []
    try:
        z = zipfile.ZipFile(ruta)
    except Exception as e:
        return [_h(ROJO, 'Archivo', f'No se pudo abrir el .pbix ({type(e).__name__}). ¿Es un archivo de Power BI Desktop?')]
    vis = _visuales(z)
    if not vis:
        return [_h(AMARILLO, 'Archivo', 'No se encontró el diseño del informe dentro del .pbix (formato no reconocido)')]
    # 1. textos con mes y año
    malos = []
    for pag, vid, obj in vis:
        acc = []; _strings(obj, acc)
        for t in acc:
            for m in RX_MES.finditer(t):
                if (MESES.index(m.group(1).lower()) + 1, int(m.group(2))) != (mes, anio):
                    malos.append((pag, t.strip("'")))
    for pag, t in sorted(set(malos)):
        H.append(_h(ROJO, f'Texto · {pag}', f'Dice "{t[:90]}" y el informe es de {MESES[mes - 1]} {anio}. Corrija el título o texto en Power BI Desktop'))
    # 2. umbrales del semáforo de avance
    esperado = sorted({round(mes / 12, 3), round((mes - 1) / 12, 3)})
    for pag, vid, obj in vis:
        acc = []; _umbrales(obj, acc)
        otros = sorted({round(x, 3) for _, x in acc if not any(abs(x - e) <= 0.006 for e in esperado)})
        if otros:
            H.append(_h(AMARILLO, f'Formato condicional · {pag}', f'El semáforo de "% de avance" usa {", ".join(f"{x:.1%}".replace(".", ",") for x in otros)}; '
                        f'para {MESES[mes - 1]} corresponde {", ".join(f"{x:.1%}".replace(".", ",") for x in esperado)} ({mes - 1}/12 y {mes}/12). '
                        'Ajústelo en Formato condicional > Iconos de la tabla'))
    # 3. valores fijos de los filtros presentes en los datos
    if base is not None:
        datos = {'Base Consolidada': base, 'Base Glosas': bg}
        vistos = set()
        for pag, vid, obj in vis:
            acc = []; _filtros_in(obj, acc, {})
            for ent, col, lits in acc:
                df = datos.get(ent)
                if df is None or col not in df.columns: continue
                presentes = set(df[col].dropna().astype(str))
                for v in lits:
                    if v not in presentes and (pag, col, v) not in vistos:
                        vistos.add((pag, col, v))
                        H.append(_h(ROJO, f'Filtro · {pag}', f'Un visual filtra "{col}" = "{v}" y ese valor no está en el consolidado: el visual queda vacío o incompleto. '
                                    'Revise el nombre en el catálogo o actualice el filtro'))
    if not any(h['Criticidad'] != VERDE for h in H):
        H.append(_h(VERDE, 'Informe Power BI', f'Títulos de {MESES[mes - 1]} {anio}, umbrales del semáforo y valores de los filtros conformes'))
    H.append(_h(AMARILLO, 'Origen de datos (revisión manual)', 'En Power BI Desktop: Transformar datos > Configuración de origen de datos. Debe apuntar al consolidado de este mes '
                'y luego presionar Actualizar. El agente no puede leer esa ruta dentro del .pbix'))
    return H
