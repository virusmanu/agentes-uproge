"""Agrega al Catálogo Maestro lo que el agente encontró sin registrar (programas, regiones, nombres del clasificador,
equivalencias de ingresos). Solo agrega filas al final de cada hoja, con el mismo formato; no modifica las existentes.
Antes de guardar deja una copia del catálogo en respaldos/catálogo."""
import os, copy, shutil, datetime
import pandas as pd

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REGIONES = ['Arica y Parinacota', 'Tarapacá', 'Antofagasta', 'Atacama', 'Coquimbo', 'Valparaíso', 'Metropolitana', "O'Higgins",
            'Maule', 'Ñuble', 'Biobío', 'Araucanía', 'Los Ríos', 'Los Lagos', 'Aysén', 'Magallanes', 'Nivel Central']
TRATAMIENTOS = ['Incluir', 'Sumar a P01', 'Excluir']


def opciones(catalogo):
    """Datos para armar la ventana: programas por servicio y tipo de llave habitual por servicio y hoja."""
    x = pd.read_excel(catalogo, sheet_name=['Programas', 'Regiones'], dtype=str)
    p = x['Programas'].fillna('')
    progs = {s: sorted({v for v in g.Programa if v}) for s, g in p.groupby('Servicio')}
    def llave(df, s, preferir=None):
        g = df[df.Servicio == s]['Tipo de llave'].dropna()
        if preferir:
            pr = [v for v in g if preferir in v]
            if pr: return pr[-1]
        return g.iloc[-1] if len(g) else ''
    return {'programas': progs,
            'llave_programa': {s: llave(p, s, 'ERP') for s in p.Servicio.unique()},
            'llave_region': {s: llave(x['Regiones'], s) for s in x['Regiones'].Servicio.unique()},
            'llave_region_cf': {s: llave(x['Regiones'], s, 'código C.F.') for s in x['Regiones'].Servicio.unique()}}


def fila_para(p, valores, op):
    """(hoja, {columna: valor}) para un pendiente completado por el usuario."""
    s = p['servicio']
    if p['tipo'] == 'Programa':
        return 'Programas', {'Servicio': s, 'Tipo de llave': op['llave_programa'].get(s, ''), 'Valor en el informe bruto': p['valor'],
                             'Programa': valores.get('programa', '') if valores.get('tratamiento') != 'Excluir' else valores.get('programa', ''),
                             'Tratamiento': valores.get('tratamiento', 'Incluir'),
                             'Nota': f'Agregado desde el agente el {datetime.date.today():%d-%m-%Y}'}
    if p['tipo'] in ('Región', 'Región CF'):
        reg = valores['region']; cod = REGIONES.index(reg) + 1
        ll = op['llave_region_cf' if p['tipo'] == 'Región CF' else 'llave_region'].get(s, '')
        return 'Regiones', {'Servicio': s, 'Tipo de llave': ll, 'Valor en el informe bruto': p['valor'], 'Código región': cod, 'Región': reg}
    if p['tipo'] == 'Nombre':
        return 'Clasificador', {'Nivel': p['nivel'], 'Código': str(p['valor']), 'Nombre': valores['nombre']}
    if p['tipo'] == 'Ingreso':
        return 'Ingresos SIGFIN', {'Servicio': s, 'Catálogo en el informe bruto': str(p['valor']), 'Código clasificador': str(valores['codigo'])}
    raise ValueError(p['tipo'])


def agregar(catalogo, filas):
    """filas: [(hoja, {columna: valor})]. Devuelve la ruta del respaldo."""
    import openpyxl
    d = os.path.join(RAIZ, 'respaldos', 'catálogo'); os.makedirs(d, exist_ok=True)
    resp = os.path.join(d, f'{datetime.datetime.now():%Y-%m-%d %H%M%S} {os.path.basename(catalogo)}')
    shutil.copy2(catalogo, resp)
    wb = openpyxl.load_workbook(catalogo)
    for hoja, valores in filas:
        ws = wb[hoja]
        cols = {str(c.value).strip(): c.column for c in ws[1] if c.value}
        ult = ws.max_row
        while ult > 1 and all(ws.cell(ult, j).value in (None, '') for j in cols.values()): ult -= 1
        n = ult + 1
        for col, v in valores.items():
            if col not in cols: continue
            j = cols[col]; c = ws.cell(n, j, v)
            c._style = copy.copy(ws.cell(ult, j)._style)
        if ws.auto_filter.ref:
            a = ws.auto_filter.ref.split(':'); ws.auto_filter.ref = f'{a[0]}:{"".join(ch for ch in a[-1] if ch.isalpha())}{n}'
    try:
        wb.save(catalogo)
    except PermissionError:
        raise PermissionError('No pude guardar el Catálogo Maestro: está abierto en Excel (o lo está usando otra persona). Ciérrelo e intente de nuevo.')
    return resp


def describir(p):
    t = {'Programa': 'Programa sin registrar', 'Región': 'Unidad sin región', 'Región CF': 'Centro financiero sin región',
         'Nombre': f'{p.get("nivel", "")} sin nombre', 'Ingreso': 'Catálogo de ingresos sin equivalencia'}[p['tipo']]
    return f'{t} · {p["servicio"]} {p["clase"]}: {p["valor"]}' + (f' ({p["nota"]})' if p.get('nota') else '')
