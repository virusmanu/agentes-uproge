"""Documentos Word del mes: borrador de Ideas Fuerza y resumen ejecutivo para la jefatura.
Se guardan en la carpeta MINAGRI del mes; no modifican el consolidado ni ningún archivo que lea Power BI."""
import os, datetime
import pandas as pd
import word

MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
NOMBRES = {'01 SUBSE': 'Subsecretaría de Agricultura', '02 ODEPA': 'Oficina de Estudios y Políticas Agrarias (ODEPA)',
           '03 INDAP': 'Instituto de Desarrollo Agropecuario (INDAP)', '04 SAG': 'Servicio Agrícola y Ganadero (SAG)',
           '05 CONAF': 'Corporación Nacional Forestal (CONAF)', '06 CNR': 'Comisión Nacional de Riego (CNR)'}
CORTO = {'01 SUBSE': 'Subsecretaría', '02 ODEPA': 'ODEPA', '03 INDAP': 'INDAP', '04 SAG': 'SAG', '05 CONAF': 'CONAF', '06 CNR': 'CNR'}
# transferencias entre servicios del Ministerio (para el total neto): SUBSE hacia INDAP y SAG; CNR hacia INDAP
INTRAPARTIDA = {'01 SUBSE': ('2402301', '2402302', '2402303', '2402304', '3302130', '3302131', '3302132', '3302133'), '06 CNR': ('3302004',)}
AVISO = 'Borrador generado por el agente. Revise y complete antes de enviar.'


def _mm(v):   # miles de $ -> millones de $, sin decimales
    return f'{v / 1000:,.0f}'.replace(',', '.')

def _p(x):
    return '—' if x is None or pd.isna(x) else f'{x * 100:,.1f}'.replace('.', ',') + '%'

def ruta(r, que):
    nombre = {'ideas': f'{r["mes"]:02d} {MESES[r["mes"] - 1]} {r["anio"]} - Ideas Fuerza (borrador agente).docx',
              'resumen': f'MINAGRI Resumen ejecutivo a {MESES[r["mes"] - 1]} {r["anio"]}.docx'}[que]
    return os.path.join(r['carpeta'], 'MINAGRI', nombre)


def ideas_fuerza(r):
    ideas = r['ideas']
    d = word.Documento(f'Ideas Fuerza – Ejecución presupuestaria a {MESES[r["mes"] - 1]} {r["anio"]}')
    d.parrafo(AVISO + ' La narrativa ministerial se complementa con el avance físico (coberturas, programas) que informan los analistas.', i=True, color='7F7F7F', tam=9)
    for k, (_, f) in enumerate(ideas.iterrows(), 1):
        inst = str(f['Institución'])
        titulo = 'Ministerio de Agricultura' if inst.upper() == 'MINAGRI' else next((v for s, v in NOMBRES.items() if CORTO[s].upper() == inst.upper() or s.endswith(inst.upper())), inst)
        d.seccion(f'{k}. {titulo}')
        for p in str(f['Ideas Fuerza'] or '').split('\n'):
            if p.strip(): d.parrafo(p.strip())
    return d.guardar(ruta(r, 'ideas'))


def _tabla_servicios(base, prev):
    g = base[base.Tipo == 'Gastos'].groupby('Servicio')[['Presupuesto Vigente', 'Monto Devengo']].sum()
    gp = prev[prev.Tipo == 'Gastos'].groupby('Servicio')['Monto Devengo'].sum() if prev is not None else None
    a = base['Asignación'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(7)
    ip = base[(base.Tipo == 'Gastos') & pd.Series([a_ in INTRAPARTIDA.get(s, ()) for s, a_ in zip(base.Servicio, a)], index=base.index)]
    intra = ip[['Presupuesto Vigente', 'Monto Devengo']].sum()
    return g, gp, intra


def resumen_ejecutivo(r):
    base, prev, anual, dec = r['base'], r.get('prev'), r.get('anual') or {}, r.get('dec') or {}
    mes, anio = r['mes'], r['anio']
    g, gp, intra = _tabla_servicios(base, prev)
    d = word.Documento(f'Resumen ejecutivo – Ejecución presupuestaria MINAGRI a {MESES[mes - 1]} {anio}')
    d.parrafo(AVISO, i=True, color='7F7F7F', tam=9)
    tv, td = g['Presupuesto Vigente'].sum(), g['Monto Devengo'].sum()
    d.seccion('1. Ejecución de gastos')
    d.parrafo(f'Al cierre de {MESES[mes - 1].lower()} de {anio}, el Ministerio de Agricultura registra un gasto devengado de ${_mm(td)} millones, '
              f'equivalente al {_p(td / tv)} de su presupuesto vigente de ${_mm(tv)} millones'
              + (f' (en igual mes de {anio - 1}: {_p(anual.get("MINAGRI"))}).' if anual.get('MINAGRI') is not None else '.'))
    filas = []
    for s in g.index:
        v, dv = g.at[s, 'Presupuesto Vigente'], g.at[s, 'Monto Devengo']
        dm = dv - gp.get(s, 0) if gp is not None else None
        filas.append([CORTO.get(s, s), _mm(v), _mm(dv), _p(dv / v if v else None), _p(anual.get(s)), _mm(dm) if dm is not None else '—'])
    dmt = td - gp.sum() if gp is not None else None
    filas.append(['MINAGRI bruto', _mm(tv), _mm(td), _p(td / tv), _p(anual.get('MINAGRI')), _mm(dmt) if dmt is not None else '—'])
    if intra['Presupuesto Vigente']:
        nv, nd = tv - intra['Presupuesto Vigente'], td - intra['Monto Devengo']
        filas.append(['MINAGRI neto (sin transferencias intrapartida)', _mm(nv), _mm(nd), _p(nd / nv), '—', '—'])
    d.tabla(['Servicio', 'Vigente (MM$)', 'Devengo (MM$)', '% avance', f'% igual mes {anio - 1}', 'Devengo del mes (MM$)'],
            filas, anchos=[2560, 1300, 1300, 1100, 1500, 1600], total=False)
    # decretos del período
    d.seccion('2. Modificaciones del presupuesto vigente en el mes')
    if dec:
        for s in sorted(dec):
            for n, desc, monto in dec[s]:
                d.vineta(f'{CORTO.get(s, s)}: Decreto N°{n}' + (f' ({desc})' if desc else '') + f', {"+" if monto >= 0 else "−"}${_mm(abs(monto))} millones.')
    else:
        d.parrafo('Sin decretos de modificación aplicados entre el corte del mes anterior y el de este mes (o sin mes anterior para comparar).')
    # revisión
    d.seccion('3. Revisión del consolidado')
    rev = r.get('revision')
    if rev is None or len(rev) == 0:
        d.parrafo('La revisión del consolidado no registra incidencias.')
    else:
        nr, na = int((rev.Nivel == 'Rojo').sum()), int((rev.Nivel == 'Amarillo').sum())
        pend = int(((rev.Nivel == 'Rojo') & (rev.Estado != 'Aceptada')).sum())
        d.parrafo(f'Incidencias: {nr} roja(s) y {na} amarilla(s); ' + (f'{pend} roja(s) pendiente(s) de explicación.' if pend else 'todas las rojas están explicadas.'))
        filas = [[x.Control, f'{x.Servicio} {x.Código}'.strip(), x.Estado + (f': {x.Comentario}' if str(x.Comentario).strip() and str(x.Comentario) != 'nan' else '')]
                 for x in rev.itertuples() if x.Nivel == 'Rojo' or x.Estado == 'Aceptada']
        if filas: d.tabla(['Control', 'Servicio y código', 'Estado y explicación'], filas, anchos=[3000, 2200, 4160], alinear=['left'] * 3)
    d.seccion('4. Fuente')
    d.parrafo(f'Consolidado del agente ({os.path.basename(r["salida"])}), fecha de corte de decretos {pd.Timestamp(r["corte"]):%d-%m-%Y}, '
              f'decretos aplicados desde su toma de razón. Cifras de gasto (devengado) en millones de pesos. '
              f'Generado el {datetime.datetime.now():%d-%m-%Y %H:%M} con la versión {r.get("version", "")} del agente.', tam=9)
    return d.guardar(ruta(r, 'resumen'))
