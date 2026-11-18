"""Verificación de la firma digital de las actualizaciones (Ed25519, RFC 8032), sin bibliotecas externas.

Cada versión publicada trae un archivo firma.json con la huella SHA-256 de cada archivo y una firma hecha con la clave
privada de UPROGE (que no está en el repositorio). El agente solo instala desde GitHub una versión cuya firma sea válida
con la clave pública de abajo y cuyos archivos coincidan con sus huellas: si alguien modificara el repositorio, la
actualización se rechaza.
"""
import hashlib, json, os

CLAVE_PUBLICA = 'daff9cb646783404877a5a462fbdfb2f3879e5bd50ab74b83ef8b60bc079b1e8'

# ---------------------------------------------------------------- Ed25519 (implementación de referencia, RFC 8032 §6)
_p = 2 ** 255 - 19
_L = 2 ** 252 + 27742317777372353535851937790883648493
_d = -121665 * pow(121666, _p - 2, _p) % _p
_I = pow(2, (_p - 1) // 4, _p)


def _inv(x): return pow(x, _p - 2, _p)


def _xrec(y):
    xx = (y * y - 1) * _inv(_d * y * y + 1)
    x = pow(xx, (_p + 3) // 8, _p)
    if (x * x - xx) % _p != 0: x = (x * _I) % _p
    if x % 2 != 0: x = _p - x
    return x


_By = 4 * _inv(5) % _p
_B = (_xrec(_By) % _p, _By % _p, 1, _xrec(_By) * _By % _p)


def _add(P, Q):
    A = (P[1] - P[0]) * (Q[1] - Q[0]) % _p; B = (P[1] + P[0]) * (Q[1] + Q[0]) % _p
    C = 2 * P[3] * Q[3] * _d % _p; D = 2 * P[2] * Q[2] % _p
    E, F, G, H = B - A, D - C, D + C, B + A
    return (E * F % _p, G * H % _p, F * G % _p, E * H % _p)


def _mul(s, P):
    Q = (0, 1, 1, 0)
    while s > 0:
        if s & 1: Q = _add(Q, P)
        P = _add(P, P); s >>= 1
    return Q


def _igual(P, Q):
    return (P[0] * Q[2] - Q[0] * P[2]) % _p == 0 and (P[1] * Q[2] - Q[1] * P[2]) % _p == 0


def _descomprimir(s):
    if len(s) != 32: return None
    y = int.from_bytes(s, 'little'); signo = y >> 255; y &= (1 << 255) - 1
    if y >= _p: return None
    x = _xrec(y)
    if (x * x * (_d * y * y + 1) - (y * y - 1)) % _p != 0: return None
    if (x & 1) != signo: x = _p - x
    return (x, y, 1, x * y % _p)


def verificar_ed25519(clave_publica, mensaje, firma):
    if len(clave_publica) != 32 or len(firma) != 64: return False
    A = _descomprimir(clave_publica); R = _descomprimir(firma[:32])
    if A is None or R is None: return False
    s = int.from_bytes(firma[32:], 'little')
    if s >= _L: return False
    h = int.from_bytes(hashlib.sha512(firma[:32] + clave_publica + mensaje).digest(), 'little') % _L
    return _igual(_mul(s, _B), _add(R, _mul(h, A)))


# ---------------------------------------------------------------- manifiesto firmado
NOMBRE = 'firma.json'
IGNORAR = {'README.md', '.gitignore', NOMBRE}       # archivos de la raíz del repositorio que no son del agente


def mensaje(manifiesto):
    return json.dumps({'version': manifiesto['version'], 'archivos': manifiesto['archivos']}, sort_keys=True, ensure_ascii=False, separators=(',', ':')).encode('utf-8')


def sha256(ruta):
    h = hashlib.sha256()
    with open(ruta, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''): h.update(b)
    return h.hexdigest()


def verificar_carpeta(carpeta):
    """Verifica la carpeta de una versión descomprimida. Devuelve (estado, detalle); estado: 'valida', 'sin firma' o 'invalida'."""
    ruta = os.path.join(carpeta, NOMBRE)
    if not os.path.exists(ruta): return 'sin firma', 'La versión no trae firma digital (firma.json).'
    try:
        m = json.load(open(ruta, encoding='utf-8'))
        ok = verificar_ed25519(bytes.fromhex(CLAVE_PUBLICA), mensaje(m), bytes.fromhex(m['firma']))
    except Exception as e:
        return 'invalida', f'La firma no se pudo leer ({type(e).__name__}).'
    if not ok: return 'invalida', 'La firma digital no corresponde a la clave de UPROGE.'
    vistos = set()
    for raiz, _, archivos in os.walk(carpeta):
        for a in archivos:
            rel = os.path.relpath(os.path.join(raiz, a), carpeta).replace('\\', '/')
            if rel in IGNORAR or '__pycache__' in rel: continue
            vistos.add(rel)
            if rel not in m['archivos']: return 'invalida', f'El archivo "{rel}" no está en la lista firmada.'
            if sha256(os.path.join(raiz, a)) != m['archivos'][rel]: return 'invalida', f'El archivo "{rel}" fue modificado después de firmar la versión.'
    faltan = sorted(set(m['archivos']) - vistos)
    if faltan: return 'invalida', f'Faltan archivos de la versión firmada (p. ej. "{faltan[0]}").'
    return 'valida', f'Firma digital válida (versión {m["version"]}, {len(m["archivos"])} archivos).'
