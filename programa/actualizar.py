"""Actualizar el agente desde la consola.

  5_ACTUALIZAR_AGENTE.bat                     busca en la carpeta compartida de actualizaciones (config.ini)
                                              o pide la ruta del .zip
  arrastrar el .zip sobre 5_ACTUALIZAR_AGENTE.bat   instala ese archivo
  5_ACTUALIZAR_AGENTE.bat --volver            vuelve a una versión respaldada"""
import sys, os, warnings
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
warnings.filterwarnings('ignore')
import actualizador as A

def pausa(sin):
    if not sin: input('\nPresione Enter para cerrar...')

def opcion(clave):
    try:
        import configparser
        c = configparser.ConfigParser(); c.read(os.path.join(A.RAIZ, 'config.ini'), encoding='utf-8')
        return c['RUTAS'].get(clave, '').strip()
    except Exception:
        return ''

def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    sin = '--sin-pausa' in sys.argv
    v = A.leer_version()
    print('=' * 70); print(f' ACTUALIZAR AGENTE DE EJECUCIÓN MINAGRI · versión instalada {v["version"]} ({v.get("fecha", "")})'); print('=' * 70)

    if '--volver' in sys.argv:
        rs = A.respaldos()
        if not rs: print('No hay respaldos guardados.'); return pausa(sin)
        for i, (p, ver, f) in enumerate(rs, 1): print(f'  {i}. versión {ver}  (respaldada el {f})')
        op = input('\nNúmero de la versión a la que quiere volver (Enter para cancelar): ').strip()
        if not op.isdigit() or not 1 <= int(op) <= len(rs): print('Cancelado.'); return pausa(sin)
        A.restaurar(rs[int(op) - 1][0]); return pausa(sin)

    if args:
        info = A.inspeccionar(args[0].strip('"'))
    else:
        info = A.buscar_en_carpeta(opcion('carpeta_actualizaciones'), v['version'])
        if info:
            print(f'Hay una versión nueva en la carpeta compartida: {os.path.basename(info["archivo"])}')
        else:
            repo = A.repositorio_github(opcion('repositorio_github'))
            if repo:
                print(f'Buscando en GitHub ({repo})...')
                try: info = A.buscar_en_github(repo, v['version'])
                except Exception as e: print(f'  No se pudo consultar GitHub: {e}')
                if info: print(f'Hay una versión nueva en GitHub: {info["version"]}')
                else: print('  Sin versiones nuevas en GitHub.')
        if not info:
            ruta = input('Arrastre aquí el .zip de la actualización (o escriba su ruta) y presione Enter:\n> ').strip().strip('"')
            if not ruta: print('Cancelado.'); return pausa(sin)
            info = A.inspeccionar(ruta)
    print(f'\nVersión del archivo: {info["version"]} ({info.get("fecha", "")})')
    for c in info.get('cambios', []): print(f'  · {c}')
    if not A.es_mas_nueva(info['version'], v['version']):
        print(f'\nATENCIÓN: esta versión NO es más nueva que la instalada ({v["version"]}).')
    if input('\n¿Instalar? (S/N): ').strip().lower() not in ('s', 'si', 'sí'):
        print('Cancelado.'); return pausa(sin)
    if info.get('url'):
        print('Descargando...'); info['archivo'] = A.descargar(info)
    r = A.instalar(info['archivo'], origen='github' if info.get('url') else 'archivo',
                   confirmar_sin_firma=lambda d: input(f'\n{d} Instálela solo si el .zip se lo entregó UPROGE. ¿Instalar igual? (S/N): ').strip().lower() in ('s', 'si', 'sí'))
    print(f'\nListo: versión {r["version"]} instalada. Si quiere deshacerla: 5_ACTUALIZAR_AGENTE.bat --volver')
    pausa(sin)

if __name__ == '__main__':
    try:
        main()
    except A.ErrorActualizacion as e:
        print(f'\nNO SE PUDO ACTUALIZAR: {e}'); pausa('--sin-pausa' in sys.argv)
    except Exception as e:
        import traceback; traceback.print_exc()
        print(f'\nError inesperado: {e}. No se modificó config.ini ni el Catálogo; la versión anterior quedó en la carpeta "respaldos".')
        pausa('--sin-pausa' in sys.argv)
