"""Bitácora del agente: registro de cada generación y de las acciones de la revisión.

Archivo "Bitácora del agente <año>.xlsx" en la carpeta de informes mensuales (la que contiene "01 Enero <año>", …),
con tres hojas:
  - Generaciones: quién, cuándo, con qué versión, corte y criterio; huella del archivo generado y resumen de alertas.
  - Insumos: cada archivo usado en una generación, con su fecha, tamaño y huella digital (SHA-256), que permite
    comprobar después que no cambió.
  - Revisión: alertas aceptadas o reabiertas y hojas de control exportadas.
No es leído por Power BI. Si el archivo está abierto o no se puede escribir, la generación sigue y queda un aviso."""
import os, hashlib, getpass, platform, datetime
import pandas as pd

MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
HOJAS = {'Generaciones': ['ID', 'Fecha', 'Usuario', 'Equipo', 'Versión del agente', 'Mes', 'Fecha de corte', 'Criterio de decretos',
                          'Archivo generado', 'Huella del archivo', 'Rojos', 'Amarillos', 'Observación'],
         'Insumos': ['ID', 'Tipo', 'Archivo', 'Carpeta', 'Modificado', 'Tamaño (bytes)', 'Huella SHA-256'],
         'Revisión': ['Fecha', 'Usuario', 'Mes', 'Acción', 'Control', 'Servicio', 'Código', 'Comentario'],
         'Publicaciones': ['Fecha', 'Mes', 'Generado por', 'Aprobado por', 'Archivo publicado', 'Huella del archivo', 'Huella de datos', 'Observación']}
CADENA = 'Huella encadenada'   # desde la 2.0: cada fila lleva la huella de su contenido y de la fila anterior


def ruta(carpeta_informes, anio):
    return os.path.join(carpeta_informes, f'Bitácora del agente {anio}.xlsx')

def huella(path, bloque=1 << 20):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for b in iter(lambda: f.read(bloque), b''): h.update(b)
    return h.hexdigest()

def _usuario():
    try: return getpass.getuser()
    except Exception: return ''

def _agregar(path, hoja, filas):
    """Agrega filas a una hoja (crea el archivo con formato si no existe)."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    if os.path.exists(path):
        wb = openpyxl.load_workbook(path)
    else:
        wb = openpyxl.Workbook(); wb.remove(wb.active)
    for h, cols in HOJAS.items():
        if h not in wb.sheetnames:
            ws = wb.create_sheet(h); ws.append(cols + [CADENA])
            for c in ws[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E5F')
            for i, c in enumerate(cols + [CADENA]): ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = max(12, min(48, len(c) + 6))
            ws.freeze_panes = 'A2'
        else:
            ws = wb[h]
            if ws.cell(1, len(cols) + 1).value != CADENA:          # bitácora anterior a la 2.0: se agrega la columna
                c = ws.cell(1, len(cols) + 1, CADENA); c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E5F')
    ws = wb[hoja]; n = len(HOJAS[hoja])
    prev = next((str(ws.cell(r, n + 1).value) for r in range(ws.max_row, 1, -1) if ws.cell(r, n + 1).value), 'inicio')
    for f in filas:
        vals = [f.get(c, '') for c in HOJAS[hoja]]
        prev = _eslabon(prev, vals); ws.append(vals + [prev])
    wb.save(path)

def _eslabon(prev, vals):
    return hashlib.sha256((str(prev) + '|' + '|'.join('' if v is None else str(v) for v in vals)).encode('utf-8')).hexdigest()

def verificar_cadena(carpeta_informes, anio):
    """Recalcula la huella encadenada de cada hoja. Devuelve [(hoja, fila, texto)] con las filas alteradas o eliminadas
    (lista vacía: la bitácora está íntegra desde que lleva huellas)."""
    import openpyxl
    p = ruta(carpeta_informes, anio)
    if not os.path.exists(p): return []
    wb = openpyxl.load_workbook(p, read_only=True); out = []
    for h, cols in HOJAS.items():
        if h not in wb.sheetnames: continue
        n = len(cols); prev = None; antiguas = False
        for i, fila in enumerate(wb[h].iter_rows(min_row=2, values_only=True), start=2):
            fila = list(fila) + [None] * (n + 1 - len(fila))
            if all(v in (None, '') for v in fila): continue
            if not fila[n]:
                if prev is not None: out.append((h, i, 'fila sin huella después de filas con huella (agregada a mano)'))
                else: antiguas = True                        # filas anteriores a la 2.0
                continue
            esperado = _eslabon(prev or 'inicio', fila[:n])
            if prev is None and antiguas:
                prev = fila[n]; continue                     # primera fila con huella después de las antiguas: punto de partida
            if fila[n] != esperado: out.append((h, i, 'el contenido de la fila o de una anterior fue modificado o eliminado'))
            prev = fila[n]
    return out

def registrar_generacion(carpeta_informes, anio, mes, r, insumos, criterio, obs=''):
    """insumos: [(tipo, ruta)]. Devuelve el ID de la generación."""
    ahora = datetime.datetime.now(); gid = ahora.strftime('%Y%m%d-%H%M%S')
    rev = r.get('revision')
    nr = int((rev.Nivel == 'Rojo').sum()) if rev is not None and len(rev) else 0
    na = int((rev.Nivel == 'Amarillo').sum()) if rev is not None and len(rev) else 0
    gen = {'ID': gid, 'Fecha': ahora.strftime('%d-%m-%Y %H:%M:%S'), 'Usuario': _usuario(), 'Equipo': platform.node(),
           'Versión del agente': r.get('version', ''), 'Mes': f'{MESES[mes - 1]} {anio}', 'Fecha de corte': f'{pd.Timestamp(r["corte"]):%d-%m-%Y}',
           'Criterio de decretos': criterio, 'Archivo generado': os.path.basename(r['salida']), 'Huella del archivo': huella(r['salida']),
           'Rojos': nr, 'Amarillos': na, 'Observación': obs}
    filas = []
    for tipo, p in insumos:
        if not p or not os.path.exists(p): continue
        st = os.stat(p)
        filas.append({'ID': gid, 'Tipo': tipo, 'Archivo': os.path.basename(p), 'Carpeta': os.path.dirname(p),
                      'Modificado': datetime.datetime.fromtimestamp(st.st_mtime).strftime('%d-%m-%Y %H:%M:%S'),
                      'Tamaño (bytes)': st.st_size, 'Huella SHA-256': huella(p)})
    p = ruta(carpeta_informes, anio)
    _agregar(p, 'Generaciones', [gen]); _agregar(p, 'Insumos', filas)
    return gid

def registrar_revision(carpeta_informes, anio, mes, accion, filas=None, comentario=''):
    """accion: 'Alerta aceptada', 'Aceptación quitada', 'Hoja de control exportada'. filas: alertas (dict o Series)."""
    base = {'Fecha': datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S'), 'Usuario': _usuario(), 'Mes': f'{MESES[mes - 1]} {anio}', 'Acción': accion}
    out = [dict(base, Control=a.get('Control', ''), Servicio=a.get('Servicio', ''), Código=str(a.get('Código', '')),
                Comentario=a.get('Comentario', '') or comentario) for a in (filas or [])] or [dict(base, Comentario=comentario)]
    _agregar(ruta(carpeta_informes, anio), 'Revisión', out)

def registrar_accion(carpeta_informes, anio, mes, accion, comentario=''):
    registrar_revision(carpeta_informes, anio, mes, accion, None, comentario)

def registrar_publicacion(carpeta_informes, anio, mes, generado_por, archivo, huella_archivo, huella_datos, obs=''):
    _agregar(ruta(carpeta_informes, anio), 'Publicaciones', [{'Fecha': datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S'), 'Mes': f'{MESES[mes - 1]} {anio}',
             'Generado por': generado_por, 'Aprobado por': _usuario(), 'Archivo publicado': archivo, 'Huella del archivo': huella_archivo,
             'Huella de datos': huella_datos, 'Observación': obs}])

def ultima_generacion(carpeta_informes, anio, mes):
    """dict de la última fila de Generaciones del mes (usuario, huella, …) o None."""
    p = ruta(carpeta_informes, anio)
    if not os.path.exists(p): return None
    try:
        r = pd.read_excel(p, sheet_name='Generaciones', dtype=str)
        r = r[r['Mes'] == f'{MESES[mes - 1]} {anio}']
        return r.iloc[-1].to_dict() if len(r) else None
    except Exception:
        return None

def publicaciones(carpeta_informes, anio):
    p = ruta(carpeta_informes, anio)
    try: return pd.read_excel(p, sheet_name='Publicaciones', dtype=str)
    except Exception: return pd.DataFrame(columns=HOJAS['Publicaciones'])

def ultima_hoja_de_control(carpeta_informes, anio, mes):
    """(fecha, usuario) de la última hoja de control exportada para el mes, o None."""
    p = ruta(carpeta_informes, anio)
    if not os.path.exists(p): return None
    try:
        r = pd.read_excel(p, sheet_name='Revisión', dtype=str)
        r = r[(r['Mes'] == f'{MESES[mes - 1]} {anio}') & (r['Acción'] == 'Hoja de control exportada')]
        return (r['Fecha'].iloc[-1], r['Usuario'].iloc[-1]) if len(r) else None
    except Exception:
        return None
