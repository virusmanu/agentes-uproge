"""Seguridad del agente (versión 2.0): catálogo protegido y publicación del consolidado oficial con doble control.

Catálogo protegido: el agente guarda la huella del Catálogo Maestro cada vez que lo usa o lo modifica desde sus
ventanas. Si al generar encuentra que cambió por otro camino (a mano en Excel o desde otro equipo), lo informa con un
resumen de lo que cambió, respalda la versión nueva y lo anota en la bitácora. Así cualquier cambio queda trazado.

Publicación con doble control: el consolidado del agente ("… - AGENTE.xlsx") pasa a ser el consolidado oficial que lee
Power BI ("MINAGRI Consolidado a <mes> <año>.xlsx") solo cuando lo aprueba una persona distinta de la que lo generó
(criterio "doble_control" de Reglas de revisión), sin rojos pendientes y sin cambios en sus datos desde que se generó
(solo se pueden editar las Ideas Fuerza). Nunca reemplaza un consolidado oficial que no haya publicado el agente.
"""
import os, json, shutil, hashlib, datetime, getpass
import pandas as pd

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DIR_CAT = os.path.join(RAIZ, 'respaldos', 'catálogo protegido')
ESTADO = os.path.join(DIR_CAT, 'estado.json')
MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']


class ErrorPublicacion(Exception): pass


def _sha(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''): h.update(b)
    return h.hexdigest()


def usuario():
    try: return getpass.getuser()
    except Exception: return os.environ.get('USERNAME', '')


# ------------------------------------------------------------------ catálogo protegido
def _leer_estado():
    try: return json.load(open(ESTADO, encoding='utf-8'))
    except Exception: return {}


def _guardar_estado(catalogo, origen):
    os.makedirs(DIR_CAT, exist_ok=True)
    h = _sha(catalogo); copia = os.path.join(DIR_CAT, f'Catálogo {datetime.datetime.now():%Y-%m-%d %H%M%S} ({origen}).xlsx')
    shutil.copy2(catalogo, copia)
    E = _leer_estado(); E[os.path.normcase(os.path.abspath(catalogo))] = {'huella': h, 'fecha': f'{datetime.datetime.now():%d-%m-%Y %H:%M}', 'copia': copia, 'origen': origen}
    json.dump(E, open(ESTADO, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    copias = sorted(f for f in os.listdir(DIR_CAT) if f.endswith('.xlsx'))
    for f in copias[:-30]:                      # se conservan las últimas 30 versiones
        try: os.remove(os.path.join(DIR_CAT, f))
        except OSError: pass


def _diferencias(antes, ahora):
    """Resumen por hoja: filas agregadas y quitadas entre dos versiones del catálogo."""
    try:
        A = pd.read_excel(antes, sheet_name=None, dtype=str); B = pd.read_excel(ahora, sheet_name=None, dtype=str)
    except Exception:
        return 'no se pudo comparar con la versión anterior'
    out = []
    for h in sorted(set(A) | set(B)):
        if h not in A: out.append(f'hoja nueva "{h}"'); continue
        if h not in B: out.append(f'se eliminó la hoja "{h}"'); continue
        fa = set(map(tuple, A[h].fillna('').values)); fb = set(map(tuple, B[h].fillna('').values))
        mas, menos = len(fb - fa), len(fa - fb)
        if mas or menos: out.append(f'{h}: ' + ', '.join(t for t in (f'{mas} fila(s) nueva(s) o cambiada(s)' if mas else '', f'{menos} fila(s) eliminada(s) o cambiada(s)' if menos else '') if t))
    return '; '.join(out) or 'cambió el formato, no los datos'


def verificar_catalogo(catalogo):
    """None si el catálogo no cambió por fuera del agente (o es la primera vez); si cambió, el resumen del cambio."""
    if not os.path.exists(catalogo): return None
    e = _leer_estado().get(os.path.normcase(os.path.abspath(catalogo)))
    if e is None:
        _guardar_estado(catalogo, 'primer registro'); return None
    if _sha(catalogo) == e['huella']: return None
    resumen = _diferencias(e['copia'], catalogo) if os.path.exists(e.get('copia', '')) else 'sin copia anterior para comparar'
    _guardar_estado(catalogo, 'cambio detectado')
    return f'desde el {e["fecha"]} cambió fuera del agente de este equipo (a mano en Excel o desde otro equipo): {resumen}. La versión anterior quedó respaldada en respaldos\\catálogo protegido'


def catalogo_modificado_por_el_agente(catalogo, motivo='editado desde el agente'):
    """Se llama después de que el agente guarda el catálogo (fecha de corte, agregar al catálogo)."""
    try: _guardar_estado(catalogo, motivo)
    except Exception: pass


# ------------------------------------------------------------------ publicación con doble control
def huella_datos(consolidado):
    """Huella de las hojas que lee Power BI con cifras (Base Consolidada y Base Glosas), leída del archivo guardado.
    No incluye Ideas Fuerza, que se edita después de generar."""
    h = hashlib.sha256()
    for hoja in ('Base Consolidada', 'Base Glosas'):
        df = pd.read_excel(consolidado, sheet_name=hoja)
        h.update(hoja.encode('utf-8')); h.update(df.to_csv(index=False, float_format='%.6f').encode('utf-8'))
    return h.hexdigest()


def ruta_oficial(carpeta_mes, anio, mes):
    return os.path.join(carpeta_mes, 'MINAGRI', f'MINAGRI Consolidado a {MESES[mes - 1]} {anio}.xlsx')


def estado_publicacion(cfg, anio, mes, carpeta_mes, salida, revision):
    """Qué falta para publicar: dict con 'puede' (bool), 'motivos' ([texto]) y datos de la última generación."""
    import bitacora as BT, revision as RV
    motivos = []
    gen = BT.ultima_generacion(cfg['carpeta_informes'], anio, mes)
    if gen is None: motivos.append('No hay una generación registrada en la bitácora para este mes: genere el consolidado.')
    if revision is not None and RV.resumen(revision)['rojos_pendientes']:
        motivos.append(f'Quedan {RV.resumen(revision)["rojos_pendientes"]} alerta(s) en rojo sin explicar en la pestaña Revisión.')
    doble = str(RV.criterio('doble_control')).strip().lower() in ('sí', 'si', 's')
    if gen is not None and doble and str(gen.get('Usuario', '')).lower() == usuario().lower():
        motivos.append(f'Doble control: el consolidado lo generó {gen.get("Usuario")}; lo debe aprobar otra persona desde su equipo '
                       '(el criterio se ajusta en Reglas de revisión, hoja Criterios).')
    if not os.path.exists(salida): motivos.append('No se encuentra el consolidado del agente.')
    return {'puede': not motivos, 'motivos': motivos, 'generacion': gen, 'doble_control': doble}


def publicar(cfg, anio, mes, carpeta_mes, salida, revision, confirmar_reemplazo=lambda: False):
    """Copia el consolidado del agente como consolidado oficial del mes. Devuelve la ruta publicada."""
    import bitacora as BT
    e = estado_publicacion(cfg, anio, mes, carpeta_mes, salida, revision)
    if not e['puede']: raise ErrorPublicacion('\n'.join(e['motivos']))
    gen = e['generacion']; obs = str(gen.get('Observación', '') or '')
    esperado = obs.split('Huella de datos: ')[-1].strip() if 'Huella de datos: ' in obs else None
    actual = huella_datos(salida)
    if esperado and actual != esperado:
        raise ErrorPublicacion('Las cifras del consolidado cambiaron después de generarlo (solo se pueden editar las Ideas Fuerza). '
                               'Vuelva a generar el mes para publicar.')
    destino = ruta_oficial(carpeta_mes, anio, mes)
    if os.path.exists(destino):
        pub = BT.publicaciones(cfg['carpeta_informes'], anio)
        if _sha(destino) not in set(pub.get('Huella del archivo', pd.Series(dtype=str)).dropna()):
            raise ErrorPublicacion(f'Ya existe "{os.path.basename(destino)}" y no lo publicó el agente (por ejemplo, un consolidado elaborado a mano). '
                                   'El agente no lo reemplaza: muévalo o renómbrelo si corresponde reemplazarlo.')
        if not confirmar_reemplazo(): raise ErrorPublicacion('Publicación cancelada.')
    tmp = destino + '.tmp'
    shutil.copy2(salida, tmp); os.replace(tmp, destino)
    BT.registrar_publicacion(cfg['carpeta_informes'], anio, mes, gen.get('Usuario', ''), os.path.basename(destino), _sha(destino), actual,
                             'Doble control' if e['doble_control'] else 'Sin doble control (criterio en Reglas de revisión)')
    return destino
