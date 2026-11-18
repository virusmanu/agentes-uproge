"""Actualización del agente en su propia carpeta, común a la ventana y a la consola.

Una actualización es el mismo .zip que se entrega para instalar (con o sin la carpeta raíz adentro).
Al instalarla:
  1. se respalda la versión actual en  respaldos\\v<versión> <fecha hora>\\
  2. se reemplazan los archivos del programa (carpeta programa, recursos, .bat, manual, historial)
  3. se conservan SIEMPRE config.ini (se agregan solo las opciones nuevas) y el Catálogo Maestro
     (si la versión trae un catálogo nuevo, se instala conservando las fechas de corte ya registradas).
Cualquier respaldo se puede reponer con restaurar()."""
import os, re, json, shutil, zipfile, datetime, tempfile, configparser, hashlib, urllib.request

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CARPETA_RESPALDOS = 'respaldos'
CATALOGO = 'Catálogo Maestro Ejecución MINAGRI 2026.xlsx'
# lo que pertenece al programa y se reemplaza al actualizar (lo demás de la carpeta no se toca)
DEL_PROGRAMA = ('programa', 'recursos', 'version.json', 'LEEME.txt', 'Manual de usuario.pdf', 'Historial de versiones.pdf')
NO_TOCAR = ('config.ini', CATALOGO, CARPETA_RESPALDOS)


class ErrorActualizacion(Exception):
    pass

YA_OFRECIDAS = set()   # versiones ya ofrecidas en esta sesión (para no preguntar dos veces)
REVISADO = [False]     # True cuando la bienvenida ya buscó versiones nuevas en esta sesión
ENCONTRADA = [None]    # la versión nueva encontrada en esa búsqueda (para volver a mostrar el aviso)


# ------------------------------------------------------------------ GitHub
ORIGEN = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'origen.json')

def repositorio_github(valor_config=''):
    """Repositorio público de donde se descargan las versiones ("usuario/repositorio").
    Se toma de config.ini (repositorio_github) si está escrito; si no, de programa/origen.json."""
    v = (valor_config or '').strip()
    if v: return v.strip('/').replace('https://github.com/', '')
    try:
        with open(ORIGEN, encoding='utf-8') as f:
            return (json.load(f).get('repositorio_github') or '').strip()
    except Exception:
        return ''

def _get(url, timeout):
    req = urllib.request.Request(url, headers={'Accept': 'application/vnd.github+json', 'User-Agent': 'Agentes-UPROGE-actualizador'})
    return urllib.request.urlopen(req, timeout=timeout)

def buscar_en_github(repo, version_actual, timeout=6, rama='main'):
    """Versión publicada en la rama principal del repositorio (archivo version.json), si es más nueva que la
    instalada. El paquete es el .zip que GitHub arma de la rama. Sin conexión o GitHub bloqueado: excepción."""
    if not repo: return None
    with _get(f'https://raw.githubusercontent.com/{repo}/{rama}/version.json', timeout) as r:
        d = json.load(r)
    version = str(d.get('version', '')).lstrip('vV')
    if not version or not es_mas_nueva(version, version_actual):
        return None
    return {'version': version, 'fecha': d.get('fecha', ''), 'cambios': d.get('cambios', []), 'origen': 'github',
            'pagina': f'https://github.com/{repo}', 'url': f'https://github.com/{repo}/archive/refs/heads/{rama}.zip', 'sha256': ''}

def descargar(info, timeout=60):
    """Descarga el .zip de una versión de GitHub a una carpeta temporal y lo verifica. Devuelve la ruta."""
    destino = os.path.join(tempfile.mkdtemp(prefix='agente_github_'), f'Agente v{info["version"]}.zip')
    h = hashlib.sha256()
    with _get(info['url'], timeout) as r, open(destino, 'wb') as f:
        while True:
            b = r.read(1 << 16)
            if not b: break
            h.update(b); f.write(b)
    if info.get('sha256') and h.hexdigest() != info['sha256'].lower():
        raise ErrorActualizacion('El archivo descargado de GitHub no coincide con el publicado (suma de verificación distinta); no se instalará.')
    v = inspeccionar(destino)['version']
    if clave(v) != clave(info['version']):
        raise ErrorActualizacion(f'El archivo descargado dice ser la versión {v}, no la {info["version"]}; no se instalará.')
    return destino


# ------------------------------------------------------------------ versiones
def leer_version(raiz=RAIZ):
    try:
        with open(os.path.join(raiz, 'version.json'), encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {'version': '0', 'fecha': '', 'cambios': []}

def clave(v):
    """Compara versiones. Esquema UPROGE: la versión estable es mayor.menor ('1.5'); las betas acumulativas se marcan
    con una letra que baja de la j a la a ('1.5j' es la primera beta, '1.5a' la última); después viene '1.6'.
    Orden: 1.5 < 1.5j < 1.5i < ... < 1.5a < 1.6. También acepta números de tres partes ('1.0.1')."""
    v = str(v).strip().lower().lstrip('v').replace(' beta', '')
    m = re.match(r'^([\d.]+?)\.?([a-j])?$', v)
    nums, letra = (m.group(1), m.group(2)) if m else (v, None)
    out = []
    for p in nums.split('.'):
        try: out.append(int(p))
        except ValueError: out.append(0)
    out = (out + [0] * 3)[:3]
    return tuple(out) + ((ord('k') - ord(letra)) if letra else 0,)

def es_beta(v):
    return clave(v)[-1] > 0

def es_mas_nueva(a, b):
    return clave(a) > clave(b)


# ------------------------------------------------------------------ paquetes
def _nombre(m):
    """Nombre real del archivo dentro del zip (corrige tildes de zips creados sin la marca UTF-8)."""
    n = m.filename
    if not m.flag_bits & 0x800:
        try: n = n.encode('cp437').decode('utf-8')
        except (UnicodeEncodeError, UnicodeDecodeError): pass
    return n.replace('\\', '/')

def _prefijo(nombres):
    """Carpeta dentro del zip que contiene programa/proceso.py ('' si está en la raíz del zip)."""
    for n in nombres:
        n2 = n.replace('\\', '/')
        if n2.endswith('programa/proceso.py'):
            return n2[:-len('programa/proceso.py')]
    return None

def _seguro(nombre):
    n = nombre.replace('\\', '/')
    return not (n.startswith('/') or ':' in n.split('/')[0] or '..' in n.split('/'))

def inspeccionar(ruta_zip):
    """Datos de la versión que trae un .zip, sin instalar nada."""
    if not zipfile.is_zipfile(ruta_zip):
        raise ErrorActualizacion(f'"{os.path.basename(ruta_zip)}" no es un archivo .zip válido.')
    with zipfile.ZipFile(ruta_zip) as z:
        nombres = [_nombre(m) for m in z.infolist()]
        pre = _prefijo(nombres)
        if pre is None:
            raise ErrorActualizacion('El .zip no contiene el agente (falta programa\\proceso.py). Use el .zip entregado como actualización.')
        malos = [n for n in nombres if not _seguro(n)]
        if malos:
            raise ErrorActualizacion(f'El .zip contiene rutas no permitidas ({malos[0]}); no se instalará.')
        try:
            mv = next(m for m in z.infolist() if _nombre(m) == pre + 'version.json')
            info = json.loads(z.read(mv).decode('utf-8'))
        except StopIteration:
            info = {'version': '1.0.1', 'fecha': '', 'cambios': ['Versión anterior al actualizador (sin registro de cambios).']}
    info['prefijo'] = pre; info['archivo'] = ruta_zip
    return info

def buscar_en_carpeta(carpeta, version_actual):
    """El .zip más nuevo de la carpeta compartida de actualizaciones, si es más nuevo que el instalado."""
    if not carpeta or not os.path.isdir(carpeta): return None
    mejor = None
    for f in os.listdir(carpeta):
        if not f.lower().endswith('.zip') or f.startswith('~$'): continue
        try:
            info = inspeccionar(os.path.join(carpeta, f))
        except Exception:
            continue
        if es_mas_nueva(info['version'], version_actual) and (mejor is None or es_mas_nueva(info['version'], mejor['version'])):
            mejor = info
    return mejor


# ------------------------------------------------------------------ respaldo / restauración
def _copiar(origen, destino):
    if os.path.isdir(origen):
        shutil.copytree(origen, destino, dirs_exist_ok=True, ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
    elif os.path.exists(origen):
        os.makedirs(os.path.dirname(destino), exist_ok=True); shutil.copy2(origen, destino)

def _archivos_del_programa(raiz):
    fs = [n for n in DEL_PROGRAMA if os.path.exists(os.path.join(raiz, n))]
    fs += [f for f in os.listdir(raiz) if f.lower().endswith('.bat')]
    return fs

def respaldar(raiz=RAIZ, motivo='antes de actualizar'):
    v = leer_version(raiz).get('version', '0')
    nombre = f'v{v} {datetime.datetime.now():%Y-%m-%d %H%M%S}'
    dest = os.path.join(raiz, CARPETA_RESPALDOS, nombre)
    for n in _archivos_del_programa(raiz) + ['config.ini', CATALOGO]:
        _copiar(os.path.join(raiz, n), os.path.join(dest, n))
    with open(os.path.join(dest, 'respaldo.txt'), 'w', encoding='utf-8') as f:
        f.write(f'Respaldo de la versión {v}, {motivo}, {datetime.datetime.now():%d-%m-%Y %H:%M}.\n')
    return dest

def respaldos(raiz=RAIZ):
    """[(carpeta, versión, fecha legible)] del más nuevo al más antiguo."""
    base = os.path.join(raiz, CARPETA_RESPALDOS)
    if not os.path.isdir(base): return []
    out = []
    for d in os.listdir(base):
        p = os.path.join(base, d)
        if os.path.isdir(p) and os.path.exists(os.path.join(p, 'programa', 'proceso.py')):
            out.append((p, leer_version(p).get('version', d.split()[0].lstrip('v')),
                        datetime.datetime.fromtimestamp(os.path.getmtime(p)).strftime('%d-%m-%Y %H:%M')))
    return sorted(out, key=lambda x: os.path.getmtime(x[0]), reverse=True)


# ------------------------------------------------------------------ instalación
def _sincronizar_carpeta(origen, destino, borrar_sobrantes):
    """Copia archivo por archivo (sin borrar la carpeta, que puede estar en uso por la ventana abierta)."""
    nuevos = set()
    for dp, dn, fn in os.walk(origen):
        dn[:] = [d for d in dn if d != '__pycache__']
        rel = os.path.relpath(dp, origen)
        os.makedirs(os.path.join(destino, rel), exist_ok=True)
        for f in fn:
            if f.endswith('.pyc'): continue
            shutil.copy2(os.path.join(dp, f), os.path.join(destino, rel, f)); nuevos.add(os.path.normpath(os.path.join(rel, f)))
    if borrar_sobrantes and os.path.isdir(destino):
        for dp, dn, fn in os.walk(destino):
            for f in fn:
                rel = os.path.normpath(os.path.relpath(os.path.join(dp, f), destino))
                if rel not in nuevos and '__pycache__' not in rel and (f.endswith('.py') or f.endswith('.pyw') or f.endswith('.ps1')):
                    try: os.remove(os.path.join(dp, f))
                    except OSError: pass
    pc = os.path.join(destino, '__pycache__')
    if os.path.isdir(pc): shutil.rmtree(pc, ignore_errors=True)

def _fusionar_config(nuevo, actual):
    """Usa la plantilla nueva (con sus comentarios y opciones nuevas) y le pone los valores del usuario."""
    if not os.path.exists(actual):
        shutil.copy2(nuevo, actual); return []
    c = configparser.ConfigParser(); c.read(actual, encoding='utf-8')
    valores = dict(c['RUTAS']) if 'RUTAS' in c else {}
    out, agregadas, vistos = [], [], set()
    for l in open(nuevo, encoding='utf-8').read().splitlines():
        k = l.split('=', 1)[0].strip()
        if '=' in l and not l.lstrip().startswith(';'):
            vistos.add(k)
            if k in valores: out.append(f'{k} = {valores[k]}'); continue
            agregadas.append(k)
        out.append(l)
    for k, v in valores.items():             # opciones propias del usuario que la plantilla no trae
        if k not in vistos: out.append(f'{k} = {v}')
    open(actual, 'w', encoding='utf-8').write('\n'.join(out) + '\n')
    return agregadas

def _fusionar_catalogo(nuevo, actual):
    """Instala el catálogo nuevo conservando las fechas de corte registradas en el actual."""
    import openpyxl
    if not os.path.exists(actual):
        shutil.copy2(nuevo, actual); return 0
    viejo = openpyxl.load_workbook(actual, read_only=True)
    cortes = []
    if 'Corte decretos' in viejo.sheetnames:
        cortes = [r for r in viejo['Corte decretos'].iter_rows(min_row=2, values_only=True) if r and r[0]]
    viejo.close()
    wb = openpyxl.load_workbook(nuevo)
    n = 0
    if cortes and 'Corte decretos' in wb.sheetnames:
        ws = wb['Corte decretos']
        ya = {str(r[0]).strip() for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0]}
        for r in cortes:
            if str(r[0]).strip() not in ya:
                ws.append(list(r)); n += 1
    tmp = actual + '.nuevo.xlsx'
    wb.save(tmp); os.replace(tmp, actual)
    return n

def instalar(ruta_zip, raiz=RAIZ, log=print):
    """Instala la versión del .zip. Devuelve un dict con lo hecho."""
    info = inspeccionar(ruta_zip)
    actual = leer_version(raiz).get('version', '0')
    log(f'Versión instalada: {actual} · versión del archivo: {info["version"]}')
    resp = respaldar(raiz, f'antes de instalar la versión {info["version"]}')
    log(f'Respaldo creado en {os.path.relpath(resp, raiz)}')
    tmp = tempfile.mkdtemp(prefix='agente_act_')
    try:
        with zipfile.ZipFile(ruta_zip) as z:
            for m in z.infolist():
                n = _nombre(m)
                if not n.startswith(info['prefijo']) or n.endswith('/'): continue
                rel = n[len(info['prefijo']):]
                if not rel or not _seguro(rel): continue
                dest = os.path.join(tmp, *rel.split('/'))
                os.makedirs(os.path.dirname(dest), exist_ok=True)
                with z.open(m) as fz, open(dest, 'wb') as fd: shutil.copyfileobj(fz, fd)
        hechos = []
        # programa: se reemplaza completo (los .py que ya no existan se eliminan)
        _sincronizar_carpeta(os.path.join(tmp, 'programa'), os.path.join(raiz, 'programa'), borrar_sobrantes=True)
        hechos.append('Programa actualizado')
        # recursos: se agregan/reemplazan, sin borrar lo que haya puesto el usuario (p. ej. logo.png)
        if os.path.isdir(os.path.join(tmp, 'recursos')):
            _sincronizar_carpeta(os.path.join(tmp, 'recursos'), os.path.join(raiz, 'recursos'), borrar_sobrantes=False)
        for f in os.listdir(tmp):
            p = os.path.join(tmp, f)
            if os.path.isfile(p) and (f.lower().endswith('.bat') or f in DEL_PROGRAMA):
                shutil.copy2(p, os.path.join(raiz, f))
        hechos.append('Accesos .bat, manual e historial actualizados')
        if os.path.exists(os.path.join(tmp, 'config.ini')):
            nuevas = _fusionar_config(os.path.join(tmp, 'config.ini'), os.path.join(raiz, 'config.ini'))
            hechos.append('config.ini conservado' + (f' (opciones nuevas: {", ".join(nuevas)})' if nuevas else ''))
        if info.get('actualiza_catalogo') and os.path.exists(os.path.join(tmp, CATALOGO)):
            n = _fusionar_catalogo(os.path.join(tmp, CATALOGO), os.path.join(raiz, CATALOGO))
            hechos.append(f'Catálogo Maestro actualizado (se conservaron {n} fecha(s) de corte propias)')
        else:
            hechos.append('Catálogo Maestro conservado sin cambios')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    for h in hechos: log(h)
    return {'version': info['version'], 'anterior': actual, 'respaldo': resp, 'hechos': hechos, 'cambios': info.get('cambios', [])}

def restaurar(carpeta_respaldo, raiz=RAIZ, log=print):
    """Vuelve a la versión guardada en un respaldo (respaldando antes la actual)."""
    if not os.path.exists(os.path.join(carpeta_respaldo, 'programa', 'proceso.py')):
        raise ErrorActualizacion('La carpeta elegida no es un respaldo del agente.')
    v = leer_version(carpeta_respaldo).get('version', '?')
    resp = respaldar(raiz, f'antes de volver a la versión {v}')
    log(f'Respaldo de la versión actual en {os.path.relpath(resp, raiz)}')
    _sincronizar_carpeta(os.path.join(carpeta_respaldo, 'programa'), os.path.join(raiz, 'programa'), borrar_sobrantes=True)
    if os.path.isdir(os.path.join(carpeta_respaldo, 'recursos')):
        _sincronizar_carpeta(os.path.join(carpeta_respaldo, 'recursos'), os.path.join(raiz, 'recursos'), borrar_sobrantes=False)
    for f in os.listdir(carpeta_respaldo):
        p = os.path.join(carpeta_respaldo, f)
        if os.path.isfile(p) and (f.lower().endswith('.bat') or f in DEL_PROGRAMA):
            shutil.copy2(p, os.path.join(raiz, f))
    if not os.path.exists(os.path.join(carpeta_respaldo, 'version.json')):
        try: os.remove(os.path.join(raiz, 'version.json'))
        except OSError: pass
    log(f'Se volvió a la versión {v}. config.ini y Catálogo Maestro no se modificaron.')
    return v
