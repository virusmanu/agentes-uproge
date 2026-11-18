"""Agente de Ejecución Presupuestaria MINAGRI – versión de consola (2_EJECUTAR_AGENTE.bat).
Uso:  python ejecutar.py                 (pregunta el mes)
      python ejecutar.py 2026-10         (mes AAAA-MM)
      python ejecutar.py 2026-10 --corte 05-11-2026
      python ejecutar.py 2026-10 --carpeta "C:\\ruta\\a\\la\\carpeta\\del\\mes"
"""
import os, sys, argparse, traceback
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from proceso import *


def main():
    ap = argparse.ArgumentParser(); ap.add_argument('mes', nargs='?'); ap.add_argument('--corte'); ap.add_argument('--carpeta')
    ap.add_argument('--sin-pausa', action='store_true'); a = ap.parse_args()
    print('=' * 72 + f'\n  AGENTE DE EJECUCIÓN PRESUPUESTARIA MINAGRI  (versión {VERSION}, consola)\n' + '=' * 72)
    cfg = leer_config()
    for p in revisar_rutas(cfg): print('AVISO: ' + p)
    mes = a.mes or input('\nMes a procesar (AAAA-MM, por ejemplo 2026-10): ').strip()
    try: anio, m = int(mes[:4]), int(mes[5:7]); assert 1 <= m <= 12
    except Exception: raise ErrorAgente(f'Mes no válido: "{mes}". Use AAAA-MM, por ejemplo 2026-10')
    mes = f'{anio}-{m:02d}'
    est = estado_del_mes(cfg, anio, m, a.carpeta)
    print(f'\nMes: {MESES[m-1]} {anio}\nCarpeta de informes: {est["carpeta"]}\nPresupuesto Vigente: {cfg["carpeta_presupuesto_vigente"]}\nCatálogo: {cfg["catalogo"]}')
    if est.get('previa'):
        print('\nRevisión previa de los insumos:')
        for n, t in est['previa']: print(f'  {"ERROR" if n == "Error" else "Aviso"}: {t}')
    err = [t for n, t in est.get('previa', []) if n == 'Error']
    if est['problemas_archivos'] or err: raise ErrorAgente((est['problemas_archivos'] or err)[0] + '\nCorrija lo indicado y vuelva a ejecutar.')
    # fecha de corte
    if a.corte: corte = leer_fecha(a.corte); origen = 'indicada al ejecutar'
    elif est['corte_catalogo'] is not None: corte = est['corte_catalogo']; origen = 'hoja "Corte decretos" del catálogo'
    else:
        sug = None
        if est['oficial']:
            d, h, e = sugerir_corte_mes(cfg, est['oficial']); sug = d
            print(f'\nEl catálogo no tiene fecha de corte para {mes}. Comparando con el consolidado oficial existente,')
            print(f'la fecha que reproduce su Presupuesto Vigente es cualquiera entre {d:%d-%m-%Y} y {h:%d-%m-%Y}' if h is not None else f'la fecha sugerida es {d:%d-%m-%Y}', '' if e == 0 else f'(con {e} servicio(s) sin cuadrar)')
        t = '' if a.sin_pausa else input(f'Fecha de corte de decretos (DD-MM-AAAA){" [Enter = " + sug.strftime("%d-%m-%Y") + "]" if sug is not None else ""}: ').strip()
        if not t and sug is None: raise ErrorAgente(f'Falta la fecha de corte para {mes}. Agréguela en la hoja "Corte decretos" del catálogo o use --corte DD-MM-AAAA')
        corte = leer_fecha(t) if t else sug; origen = 'ingresada por el usuario' if t else 'sugerida a partir del consolidado oficial'
    print(f'Fecha de corte de decretos: {corte:%d-%m-%Y} ({origen})\n\nProcesando...')
    r = procesar(cfg, anio, m, corte, origen, est['carpeta'], log=lambda t: print('  ' + t))
    ctrl, av, val = r['ctrl'], r['avisos'], r['validacion']
    print('\n' + '-' * 72 + f'\nLISTO. Archivo generado:\n  {r["salida"]}\n' + '-' * 72)
    print(ctrl[['Servicio', 'Tipo', 'Presupuesto Vigente', 'Monto Devengo', '% Avance (Devengo/Vigente)']].to_string(index=False))
    rev = [x for x in av if x[0] == 'Revisar']
    print(f'\nAvisos: {len(av)} ({len(rev)} para revisar). Ver hoja "Control".')
    for _, t in rev: print('  * ' + t)
    if r.get('revision') is not None:
        import revision as RV
        rs = RV.resumen(r['revision'])
        print(f'\nRevisión: {rs["rojos"]} rojo(s) ({rs["rojos_pendientes"]} sin explicar) y {rs["amarillos"]} amarillo(s). Detalle en:\n  {r["ruta_revision"]}')
        print('  Para comentar las alertas y exportar la hoja de control use la versión de ventana (pestaña Revisión).')
    if r.get('semaforo') is not None:
        print('\nValidación del corte (detalle en ' + os.path.basename(r.get('ruta_validacion') or '') + '):')
        for x in r['semaforo'].itertuples(): print(f'  [{x.Estado:8}] {x.Control}: {x.Resultado}')
        for x in r.get('pbix') or []: print(f'  [{x["Criticidad"]:8}] Power BI · {x["Elemento"]}: {x["Detalle"]}')
    if val is not None:
        print('\nComparación con el consolidado oficial (hoja "Validación"):'); print(val.to_string(index=False))


if __name__ == '__main__':
    pausa = '--sin-pausa' not in sys.argv
    try: main()
    except ErrorAgente as e: print(f'\nERROR: {e}')
    except Exception: print('\nERROR INESPERADO (envíe este texto a Manuel):\n' + traceback.format_exc())
    if pausa and sys.stdin.isatty(): input('\nPresione Enter para cerrar...')
