"""Asistente de cambio de año (por ejemplo, de 2026 a 2027).

1. Crea "Catálogo Maestro Ejecución MINAGRI <año nuevo>.xlsx" a partir del catálogo vigente: conserva servicios,
   regiones, centros financieros, clasificador, ingresos y reglas; deja vacía la hoja "Corte decretos".
2. Si ya están las planillas de Presupuesto Vigente del año nuevo, compara sus programas con la hoja Programas
   del catálogo y lista los que hay que revisar.
3. Ajusta config.ini: la carpeta de informes del año que termina pasa a ser la del "año anterior" (comparación
   de Ideas Fuerza y resumen ejecutivo) y se apuntan las carpetas del año nuevo. config.ini queda respaldado.
No modifica el catálogo del año que termina ni ningún archivo de informes."""
import os, re, shutil, datetime
import pandas as pd

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def nombre_catalogo(anio):
    return f'Catálogo Maestro Ejecución MINAGRI {anio}.xlsx'


def crear_catalogo(actual, anio):
    import openpyxl
    destino = os.path.join(os.path.dirname(actual), nombre_catalogo(anio))
    if os.path.exists(destino):
        raise FileExistsError(f'Ya existe {os.path.basename(destino)}; no lo reemplazo. Si quiere rehacerlo, muévalo antes a otra carpeta.')
    m = re.search(r'(\d{4})\.xlsx$', actual); anterior = m.group(1) if m else str(anio - 1)
    wb = openpyxl.load_workbook(actual)
    ws = wb['Corte decretos']
    if ws.max_row > 1: ws.delete_rows(2, ws.max_row - 1)
    if ws.auto_filter.ref: ws.auto_filter.ref = 'A1:D1'
    if 'LEEME' in wb.sheetnames:
        for fila in wb['LEEME'].iter_rows():
            for c in fila:
                if isinstance(c.value, str) and anterior in c.value: c.value = c.value.replace(anterior, str(anio))
        wb['LEEME'].append([]); wb['LEEME'].append([f'Creado el {datetime.date.today():%d-%m-%Y} a partir del catálogo {anterior} (asistente de cambio de año). '
                                                    f'Revise las hojas Programas y Glosas contra la Ley de Presupuestos {anio}.'])
    wb.save(destino)
    return destino


def comparar_programas(catalogo, carpeta_pv):
    """Programas de las planillas de Presupuesto Vigente del año nuevo que no están en la hoja Programas, y viceversa."""
    from motor import planillas_pv
    pvs = planillas_pv(carpeta_pv)
    cat = pd.read_excel(catalogo, sheet_name='Programas', dtype=str).fillna('')
    out = []
    for s, p in pvs.items():
        b = pd.read_excel(p, sheet_name='Base Presupuesto', usecols=['Programa'], dtype=str)
        nuevos = sorted(set(b.Programa.dropna().str.strip()) - set(cat[cat.Servicio == s].Programa.str.strip()))
        sin = sorted(set(cat[(cat.Servicio == s) & (cat.Tratamiento != 'Excluir')].Programa.str.strip()) - set(b.Programa.dropna().str.strip()) - {''})
        for x in nuevos: out.append(f'{s}: programa nuevo en la planilla, no está en el catálogo: {x}')
        for x in sin: out.append(f'{s}: programa del catálogo que no aparece en la planilla del año nuevo: {x}')
    return out


def preparar(cfg, anio, carpeta_informes, carpeta_pv):
    """Devuelve el informe del cambio de año (lista de líneas)."""
    import proceso as P
    linea = []
    cat = crear_catalogo(cfg['catalogo'], anio)
    linea.append(f'✔ Catálogo creado: {cat}')
    linea.append('  Se copiaron servicios, regiones, centros financieros, clasificador, ingresos, glosas y reglas. La hoja "Corte decretos" quedó vacía.')
    if carpeta_pv and os.path.isdir(carpeta_pv):
        try:
            dif = comparar_programas(cat, carpeta_pv)
            linea += (['⚠ Programas a revisar en la hoja Programas del catálogo nuevo:'] + ['  · ' + d for d in dif]) if dif else \
                     ['✔ Los programas de las planillas del año nuevo coinciden con el catálogo.']
        except Exception as e:
            linea.append(f'⚠ No se pudieron comparar los programas con las planillas del año nuevo: {e}')
    else:
        linea.append('⚠ Aún no están las planillas de Presupuesto Vigente del año nuevo: cuando estén, revise la hoja Programas del catálogo.')
    linea.append(f'⚠ Revise también contra la Ley de Presupuestos {anio}: hoja Glosas del catálogo y, en "Reglas de revisión.xlsx", '
                 'las hojas Excedibles y Transferencias intrapartida.')
    shutil.copy2(P.CONFIG, P.CONFIG + f'.antes de {anio}')
    nuevo = dict(cfg)
    nuevo['carpeta_informes_anio_anterior'] = cfg['carpeta_informes']
    if carpeta_informes: nuevo['carpeta_informes'] = carpeta_informes
    if carpeta_pv: nuevo['carpeta_presupuesto_vigente'] = carpeta_pv
    # catálogo junto al programa: se deja vacío y se usa el del año más reciente; si es uno compartido, se apunta al nuevo
    nuevo['catalogo'] = '' if os.path.dirname(os.path.abspath(cat)) == os.path.abspath(RAIZ) else cat
    P.guardar_config(nuevo)
    linea.append(f'✔ config.ini actualizado (respaldo: config.ini.antes de {anio}): año anterior = {cfg["carpeta_informes"]}')
    return linea
