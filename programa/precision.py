"""Controles de precisión del corte (versión 1.7).

Revisan que el consolidado que lee Power BI sea fiel a los insumos y respete la estructura del informe:
  1. Contrato de entrada: los informes brutos y las planillas de Presupuesto Vigente vienen completos y legibles.
  2. Cuadratura punta a punta: lo que entra es lo que sale, al peso, por servicio y tipo (ejecución y presupuesto).
  3. Motor sombra: un segundo cálculo, por otro camino, de la ejecución por programa y subtítulo.
  4. Coherencia entre servicios: decretos aplicados en todos los servicios que nombran y transferencias
     intrapartida iguales en quien transfiere y quien recibe.
  5. Contrato de salida: columnas, tipos, nombres y valores que usan los filtros del informe Power BI.
  6. Huella del corte: con los mismos insumos, el agente produce exactamente el mismo consolidado.

Cada hallazgo tiene una criticidad: Rojo (cambia cifras o rompe el informe: debe corregirse o explicarse),
Amarillo (revisar) o Verde (control sin observaciones). Nada de esto modifica los archivos de entrada.
"""
import os, re, json, hashlib, datetime
import pandas as pd
import motor as M

ROJO, AMARILLO, VERDE = 'Rojo', 'Amarillo', 'Verde'
TOL = 0.0015            # miles de $: 1 peso, más el redondeo de las sumas
RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HUELLAS = os.path.join(RAIZ, 'respaldos', 'huellas del corte.json')

# Valores que usan los filtros del informe Power BI (si no existen en los datos, el visual queda vacío).
LITERALES = {
    ('Base Consolidada', 'Tipo'): ['Gastos', 'Ingresos'],
    ('Base Consolidada', 'Nombre Subtítulo'): ['24 TRANSFERENCIAS CORRIENTES', '32 PRESTAMOS', '33 TRANSFERENCIAS DE CAPITAL'],
    ('Base Consolidada', 'Nombre Asignación'): ['2401404 Emergencias', '2401601 Emergencias Sanitarias', '2401369 Emergencias Agrícolas'],
    ('Base Glosas', 'Glosa'): ['Viático Nacional', 'Viático Exterior', 'Horas Extraordinarias'],
}

GRUPOS = ['Contrato de entrada', 'Cuadratura punta a punta', 'Motor sombra', 'Coherencia entre servicios', 'Contrato de salida', 'Huella del corte']


def _m(v):
    d = 0 if abs(v - round(v)) < 0.0005 else 3     # miles de $: sin decimales salvo que haya pesos
    return f'{v:,.{d}f}'.replace(',', 'X').replace('.', ',').replace('X', '.')


def _h(grupo, control, crit, detalle, servicio='', revisar=''):
    return {'Grupo': grupo, 'Control': control, 'Criticidad': crit, 'Servicio': servicio, 'Detalle': detalle, 'Qué revisar': revisar}


# ------------------------------------------------------------------ 1. contrato de entrada
REQ = {('SIGFIN', 'Gastos'): ['A.U.', 'C.F', 'P.P.', 'Catalogo', 'Presupuesto', 'Compromiso', 'Devengo', 'U.E.'],
       ('SIGFIN', 'Ingresos'): ['A.U.', 'C.F', 'P.P.', 'Catalogo', 'Presupuesto', 'Devengo', 'U.E.'],
       ('SIGFE', 'Gastos'): ['Código Concepto', 'Descripción Unidad Ejecutora', 'Código Unidad Ejecutora'] + M.MONTOS,
       ('SIGFE', 'Ingresos'): ['Código Concepto', 'Descripción Unidad Ejecutora', 'Código Unidad Ejecutora'] + M.MONTOS}
NUM = {('SIGFIN', 'Gastos'): ['Presupuesto', 'Compromiso', 'Devengo'], ('SIGFIN', 'Ingresos'): ['Presupuesto', 'Devengo'],
       ('SIGFE', 'Gastos'): M.MONTOS, ('SIGFE', 'Ingresos'): M.MONTOS}


def contrato_entrada(brutos):
    H = []; G = 'Contrato de entrada'
    fmt_serv = {}
    for (s, t), b in sorted(brutos.items()):
        df, fmt, a = b['df'], b['formato'], b['archivo']
        fmt_serv.setdefault(s, set()).add(fmt)
        falta = [c for c in REQ[(fmt, t)] if c not in df.columns]
        if falta:
            H.append(_h(G, 'Columnas del informe bruto', ROJO, f'{a}: faltan las columnas {", ".join(falta)}', s,
                        'Descargue de nuevo el informe con el formato de siempre (sin quitar columnas)'))
            continue
        datos = df[pd.to_numeric(df['A.U.'], errors='coerce').fillna(0) != 0] if fmt == 'SIGFIN' else df[df['Código Concepto'].notna()]
        if not len(datos):
            H.append(_h(G, 'Informe sin datos', ROJO, f'{a}: no tiene filas de datos', s, 'Revise la descarga del informe'))
            continue
        for c in NUM[(fmt, t)]:
            v = datos[c]; malo = v.notna() & (v.astype(str).str.strip() != '') & pd.to_numeric(v, errors='coerce').isna()
            if malo.any():
                filas = ', '.join(str(i + b['df'].attrs.get('fila0', 2)) for i in datos.index[malo][:5])
                H.append(_h(G, 'Montos no numéricos', ROJO, f'{a}: {int(malo.sum())} valor(es) no numéricos en "{c}" (filas {filas}); el agente los leería como 0', s,
                            'Abra el informe y corrija esas celdas (texto en vez de número)'))
        cod = datos['Catalogo' if fmt == 'SIGFIN' else 'Código Concepto'].map(M.norm)
        malos = cod[cod.notna() & ~cod.astype(str).str.fullmatch(r'\d+')]
        if len(malos):
            H.append(_h(G, 'Códigos de clasificador no numéricos', AMARILLO, f'{a}: {len(malos)} fila(s) con código no numérico (p. ej. "{malos.iloc[0]}")', s,
                        'Esas filas quedan sin subtítulo. Revise el informe'))
        mont = NUM[(fmt, t)]
        dup = datos[datos.duplicated(keep=False) & (datos[mont].apply(pd.to_numeric, errors='coerce').fillna(0).abs().sum(axis=1) > 0)]
        if len(dup):
            H.append(_h(G, 'Filas repetidas', AMARILLO, f'{a}: {len(dup)} filas idénticas (todas las columnas iguales). Si es una descarga duplicada, la ejecución queda sumada dos veces', s,
                        'Compare con el sistema contable; si son legítimas, acepte con comentario'))
    for s, f in fmt_serv.items():
        if len(f) > 1:
            H.append(_h(G, 'Formato de informe distinto entre Gastos e Ingresos', AMARILLO, f'{s}: Gastos e Ingresos vienen de sistemas distintos ({", ".join(sorted(f))})', s,
                        'Confirme que ambos informes son del servicio y del mes correctos'))
    if not H: H.append(_h(G, 'Informes brutos', VERDE, f'{len(brutos)} informes con todas sus columnas, montos numéricos y sin filas repetidas'))
    return H


def contrato_planillas(pvs, traza_pv):
    H = []; G = 'Contrato de entrada'
    import openpyxl
    for s, ruta in sorted(pvs.items()):
        a = os.path.basename(ruta)
        try:
            wf = openpyxl.load_workbook(ruta, read_only=True)['Base Presupuesto']
            wv = openpyxl.load_workbook(ruta, read_only=True, data_only=True)['Base Presupuesto']
        except Exception as e:
            H.append(_h(G, 'Planilla de Presupuesto Vigente', ROJO, f'{a}: no se pudo leer la hoja "Base Presupuesto" ({type(e).__name__})', s)); continue
        filas_f = wf.iter_rows(values_only=True); filas_v = wv.iter_rows(values_only=True)
        cab = [str(c).strip() if c is not None else '' for c in next(filas_f)]; next(filas_v)
        if 'Monto M$' not in cab:
            H.append(_h(G, 'Planilla de Presupuesto Vigente', ROJO, f'{a}: la hoja "Base Presupuesto" no tiene la columna "Monto M$"', s)); continue
        im = cab.index('Monto M$'); it = cab.index('Tipo') if 'Tipo' in cab else None; ip = cab.index('Programa') if 'Programa' in cab else None
        sin_calc = []; sin_clase = []
        for n, (rf, rv) in enumerate(zip(filas_f, filas_v), start=2):
            f = rf[im] if im < len(rf) else None; v = rv[im] if im < len(rv) else None
            if isinstance(f, str) and f.startswith('=') and v is None: sin_calc.append(n)
            if v is not None and ((it is not None and not rv[it]) or (ip is not None and not rv[ip])): sin_clase.append(n)
        if sin_calc:
            H.append(_h(G, 'Montos con fórmula sin calcular', ROJO, f'{a}: {len(sin_calc)} fila(s) con el monto en fórmula sin valor guardado (filas {", ".join(map(str, sin_calc[:8]))}). El agente no puede leer esos montos', s,
                        'Abra la planilla en Excel, presione Ctrl+Alt+F9, guarde y vuelva a generar'))
        if sin_clase:
            H.append(_h(G, 'Montos sin programa o tipo', ROJO, f'{a}: {len(sin_clase)} fila(s) con monto pero sin Programa o Tipo (filas {", ".join(map(str, sin_clase[:8]))})', s,
                        'Complete el Programa y el Tipo (Gastos/Ingresos) en la planilla'))
    # presupuesto balanceado: gastos = ingresos por servicio (vigente del mes y ley)
    if traza_pv is not None and len(traza_pv):
        t = traza_pv[traza_pv['Incluido en el vigente'] == 'Sí']
        for s, x in t.groupby('Servicio'):
            g = x.groupby('Tipo')['Monto M$'].sum(); d = g.get('Gastos', 0) - g.get('Ingresos', 0)
            if abs(d) > TOL:
                H.append(_h(G, 'Presupuesto vigente no balanceado', ROJO, f'{s}: vigente de gastos {_m(g.get("Gastos", 0))} y de ingresos {_m(g.get("Ingresos", 0))} M$ (diferencia {_m(d)}). '
                            'Un decreto quedó registrado solo por un lado o con montos distintos', s, 'Revise en la planilla el último decreto registrado (gastos e ingresos)'))
            ley = x[x.Documento != 'Decreto'].groupby('Tipo')['Monto M$'].sum(); d = ley.get('Gastos', 0) - ley.get('Ingresos', 0)
            if abs(d) > TOL:
                H.append(_h(G, 'Presupuesto ley no balanceado', ROJO, f'{s}: ley de gastos {_m(ley.get("Gastos", 0))} y de ingresos {_m(ley.get("Ingresos", 0))} M$ (diferencia {_m(d)})', s,
                            'Revise las filas de Ley de Presupuestos en la planilla'))
    if not H: H.append(_h(G, 'Planillas de Presupuesto Vigente', VERDE, f'{len(pvs)} planillas legibles, con montos calculados y presupuesto balanceado (gastos = ingresos)'))
    return H


# ------------------------------------------------------------------ 2. cuadratura punta a punta
def cuadratura(base, cuad, traza_pv):
    H = []; G = 'Cuadratura punta a punta'
    for (s, t), c in sorted(cuad.items()):
        b = base[(base.Servicio == s) & (base.Tipo == t)]
        for m in M.MONTOS:
            if f'{m} informe' not in c: continue
            esperado = c[f'{m} informe'] - c[f'{m} excluido']; obtenido = float(b[m].sum()); d = obtenido - esperado
            if abs(d) > TOL:
                H.append(_h(G, f'Ejecución: {m}', ROJO, f'{s} {t}: el informe bruto (descontado lo excluido por programa) suma {_m(esperado)} M$ y el consolidado {_m(obtenido)} M$ (diferencia {_m(d)})', s,
                            'Hay filas que se perdieron o duplicaron al consolidar. Revise la hoja Trazabilidad ejecución'))
    if traza_pv is not None and len(traza_pv):
        t = traza_pv[traza_pv['Incluido en el vigente'] == 'Sí'].copy()
        t['Subtítulo'] = pd.to_numeric(t['Subtítulo'], errors='coerce')
        k = ['Servicio', 'Tipo', 'Programa', 'Subtítulo']
        esp = t.groupby(k, dropna=False)['Monto M$'].sum()
        bb = base[base['Presupuesto Vigente'].notna()].copy(); bb['Subtítulo'] = pd.to_numeric(bb['Subtítulo'], errors='coerce')
        obt = bb.groupby(k, dropna=False)['Presupuesto Vigente'].sum()
        j = pd.concat([esp, obt], axis=1, keys=['e', 'o']).fillna(0); j = j[(j.e - j.o).abs() > TOL]
        for (s, tp, p, st), r in j.iterrows():
            H.append(_h(G, 'Presupuesto vigente', ROJO, f'{s} {tp} {p} subtítulo {st}: la planilla (filas incluidas al corte) suma {_m(r.e)} M$ y el consolidado {_m(r.o)} M$', s,
                        'Revise la hoja Trazabilidad presupuesto'))
        ley_e = t[t.Documento != 'Decreto'].groupby(['Servicio', 'Tipo'])['Monto M$'].sum()
        ley_o = bb.groupby(['Servicio', 'Tipo'])['Presupuesto Ley'].sum()
        j = pd.concat([ley_e, ley_o], axis=1, keys=['e', 'o']).fillna(0); j = j[(j.e - j.o).abs() > TOL]
        for (s, tp), r in j.iterrows():
            H.append(_h(G, 'Presupuesto ley', ROJO, f'{s} {tp}: la planilla suma {_m(r.e)} M$ de ley y el consolidado {_m(r.o)} M$', s))
    if not H: H.append(_h(G, 'Insumos = consolidado', VERDE, 'Requerimiento, compromiso, devengo, presupuesto ley y vigente cuadran al peso con los informes y las planillas, por servicio, tipo, programa y subtítulo'))
    return H


# ------------------------------------------------------------------ 3. motor sombra
def sombra(K, brutos, base):
    """Segundo cálculo de la ejecución, directo desde los informes brutos y sin pasar por el motor: suma por programa
    (según la hoja Programas del catálogo) y subtítulo (dos primeros dígitos del código). Debe coincidir con el consolidado."""
    H = []; G = 'Motor sombra'; filas = []
    for (s, t), b in brutos.items():
        df, fmt = b['df'], b['formato']
        try:
            if fmt == 'SIGFIN':
                d = df[pd.to_numeric(df['A.U.'], errors='coerce').fillna(0) != 0]
                prog = d['P.P.'].map(M.norm); cod = d['Catalogo'].map(M.norm)
                if t == 'Ingresos': cod = cod.map(lambda v: K.ing.get((s, v)))
                mon = {'Monto Requerimiento': d['Presupuesto'], 'Monto Compromiso': d['Compromiso' if t == 'Gastos' else 'Devengo'], 'Monto Devengo': d['Devengo']}
            else:
                d = df[df['Código Concepto'].notna()]
                pc = next(c for c in d.columns if d[c].dropna().astype(str).str.startswith('ProgramaPresupuestario').mean() > 0.9)
                prog = d[pc].map(M.norm); cod = d['Código Concepto'].map(M.norm)
                mon = {m: d[m] for m in M.MONTOS}
        except Exception as e:
            H.append(_h(G, 'Cálculo independiente', AMARILLO, f'{s} {t}: no se pudo calcular ({type(e).__name__}: {e})', s)); continue
        defecto = K.defecto.get((s, t), 'Excluir') == 'Sumar a P01' and bool(K.p01.get(s))
        pr = []
        for v in prog:
            z = K.prog.get((s, v))
            if z is None: pr.append(K.p01[s] if defecto else None)       # programa nuevo: tratamiento por defecto del servicio
            else: pr.append(z[1] if z[0] in ('Incluir', 'Sumar a P01') and z[1] else None)
        x = pd.DataFrame({'Servicio': s, 'Tipo': t, 'Programa': pr, 'Subtítulo': pd.to_numeric(cod.astype(str).str[:2], errors='coerce')}, index=d.index)
        for m, v in mon.items(): x[m] = pd.to_numeric(v, errors='coerce').fillna(0) / 1000
        filas.append(x[x.Programa.notna()])
    if not filas: return H
    S = pd.concat(filas); k = ['Servicio', 'Tipo', 'Programa', 'Subtítulo']
    b = base.copy(); b['Subtítulo'] = pd.to_numeric(b['Subtítulo'], errors='coerce')
    s1 = S.groupby(k, dropna=False)[M.MONTOS].sum(); s2 = b.groupby(k, dropna=False)[M.MONTOS].sum()
    j = s1.join(s2, rsuffix=' (consolidado)', how='outer').fillna(0); n = 0
    for m in M.MONTOS:
        dd = j[(j[m] - j[m + ' (consolidado)']).abs() > TOL]
        for (s, t, p, st), r in dd.iterrows():
            n += 1
            if n <= 40:
                H.append(_h(G, f'Ejecución por programa y subtítulo: {m}', ROJO, f'{s} {t} {p} subtítulo {st if pd.notna(st) else "(sin código)"}: cálculo independiente {_m(r[m])} M$; consolidado {_m(r[m + " (consolidado)"])} M$', s,
                            'El motor y el cálculo independiente no coinciden: no use el archivo hasta aclararlo (avise a soporte del agente)'))
    if not H: H.append(_h(G, 'Cálculo independiente = consolidado', VERDE, f'{len(j)} combinaciones de servicio, tipo, programa y subtítulo coinciden al peso en requerimiento, compromiso y devengo'))
    return H


# ------------------------------------------------------------------ 4. coherencia entre servicios
def coherencia(base, traza_pv, pares):
    H = []; G = 'Coherencia entre servicios'
    if traza_pv is not None and len(traza_pv):
        d = traza_pv[(traza_pv.Documento == 'Decreto') & traza_pv['Numero Documento'].notna()]
        for n, x in d.groupby('Numero Documento'):
            inc = sorted(x.loc[x['Incluido en el vigente'] == 'Sí', 'Servicio'].unique()); exc = sorted(set(x.Servicio) - set(inc))
            if inc and exc:
                fe = '; '.join(f'{s} {f:%d-%m-%Y}' for s, f in x.groupby('Servicio')['Fecha Documento'].min().items() if pd.notna(f))
                H.append(_h(G, 'Decreto aplicado solo en algunos servicios', ROJO, f'Decreto N°{int(n)}: quedó dentro del vigente en {", ".join(inc)} y fuera en {", ".join(exc)} (fechas aplicadas: {fe})',
                            ', '.join(inc + exc), 'Un decreto entre servicios debe aplicarse con la misma fecha en todos. Revise la fecha en las planillas o regístrela en el Registro de decretos'))
    num = lambda c: pd.to_numeric(base[c], errors='coerce')
    for s1, a1, s2, p2, a2 in pares:
        x1 = base[(base.Servicio == s1) & (base.Tipo == 'Gastos') & (num('Asignación') == a1)]
        x2 = base[(base.Servicio == s2) & (base.Tipo == 'Ingresos') & (num('Asignación') == a2) & base.Programa.astype(str).str.startswith(p2 + ' ')]
        if not len(x1) and not len(x2): continue
        v1, v2 = float(x1['Presupuesto Vigente'].sum()), float(x2['Presupuesto Vigente'].sum())
        if abs(v1 - v2) > 1:
            H.append(_h(G, 'Transferencia intrapartida descuadrada', ROJO, f'{s1} transfiere {_m(v1)} M$ (gasto {a1}) y {s2} {p2} registra {_m(v2)} M$ por recibir (ingreso {a2}); diferencia {_m(v1 - v2)} M$',
                        f'{s1}, {s2}', 'Falta aplicar un decreto en uno de los dos servicios, o está con fecha distinta'))
    # transferencias de SUBSE a servicios del ministerio sin par definido en Reglas de revisión
    con_par = {a for _, a, *_ in pares}
    x = base[(base.Servicio == '01 SUBSE') & (base.Tipo == 'Gastos') & num('Asignación').isin(range(2402000, 2403000)) & (base['Presupuesto Vigente'].fillna(0) != 0)]
    for a, nom in sorted(set(zip(pd.to_numeric(x['Asignación']).astype(int), x['Nombre Asignación'].fillna('')))):
        if a not in con_par and any(k in str(nom).upper() for k in ('INDAP', 'SAG', 'CONAF', 'CNR', 'ODEPA')):
            H.append(_h(G, 'Transferencia intrapartida sin par de control', AMARILLO, f'01 SUBSE {a} {nom}: no está en la hoja "Transferencias intrapartida" de Reglas de revisión, por lo que no se controla contra el ingreso del servicio que recibe',
                        '01 SUBSE', 'Agregue el par (asignación de gasto y de ingreso) en Reglas de revisión'))
    if not H: H.append(_h(G, 'Decretos y transferencias', VERDE, 'Cada decreto quedó dentro o fuera del vigente en todos los servicios que lo registran, y las transferencias intrapartida cuadran'))
    return H


# ------------------------------------------------------------------ 5. contrato de salida
def contrato_salida(base, bg, ideas, K, anterior=None):
    H = []; G = 'Contrato de salida'
    if list(base.columns) != M.COLS:
        H.append(_h(G, 'Columnas de Base Consolidada', ROJO, 'Las columnas no son las del informe: ' + ', '.join(base.columns)))
    if list(bg.columns) != M.COLS_GLOSAS:
        H.append(_h(G, 'Columnas de Base Glosas', ROJO, 'Las columnas no son las del informe: ' + ', '.join(bg.columns)))
    if ideas is not None and list(ideas.columns)[:3] != ['Código Institución', 'Institución', 'Ideas Fuerza']:
        H.append(_h(G, 'Columnas de Ideas Fuerza', ROJO, 'Las columnas no son las del informe'))
    for df, hoja in ((base, 'Base Consolidada'), (bg, 'Base Glosas')):
        mal = ~df.Servicio.astype(str).str.fullmatch(r'\d{2} \S+')
        if mal.any(): H.append(_h(G, 'Formato de Servicio', ROJO, f'{hoja}: {int(mal.sum())} fila(s) con Servicio fuera del formato "NN SIGLA" (Power BI separa el código por el espacio)'))
    if not base.Tipo.isin(['Gastos', 'Ingresos']).all():
        H.append(_h(G, 'Valores de Tipo', ROJO, 'Hay filas con Tipo distinto de Gastos o Ingresos'))
    for c in ('Código', 'Subtítulo', 'Ítem', 'Asignación', 'Subasignación'):
        v = base[c].dropna(); malo = pd.to_numeric(v, errors='coerce').isna() | (pd.to_numeric(v, errors='coerce') % 1 != 0)
        if malo.any(): H.append(_h(G, f'Código no entero en {c}', ROJO, f'{int(malo.sum())} fila(s) con "{c}" no entero: Power BI los convierte a número entero y esas celdas quedarían con error'))
    sin_st = base['Subtítulo'].isna() | base['Nombre Subtítulo'].isna()
    if sin_st.any():
        r = base[sin_st]
        H.append(_h(G, 'Filas sin subtítulo', ROJO, f'{int(sin_st.sum())} fila(s) sin subtítulo o sin su nombre (vigente {_m(r["Presupuesto Vigente"].sum())}, devengo {_m(r["Monto Devengo"].sum())} M$): no aparecen en las tablas por subtítulo',
                    ', '.join(sorted(r.Servicio.unique())), 'Revise el clasificador del catálogo o el informe base'))
    sin_it = base['Ítem'].notna() & base['Nombre Ítem'].isna()
    if sin_it.any():
        H.append(_h(G, 'Ítems sin nombre', AMARILLO, f'{int(sin_it.sum())} fila(s) con código de ítem sin nombre: ' + ', '.join(sorted({str(int(x)) for x in base.loc[sin_it, 'Ítem']})[:10]),
                    ', '.join(sorted(base.loc[sin_it, 'Servicio'].unique())), 'Agregue el nombre en la hoja Clasificador del catálogo'))
    a = pd.to_numeric(base['Asignación'], errors='coerce')
    c000 = (a % 1000 == 0) & base['Nombre Asignación'].notna()
    if c000.any():
        H.append(_h(G, 'Asignación "000" con nombre', ROJO, f'{int(c000.sum())} fila(s): la ejecución de un ítem sin asignaciones aparecería en una línea aparte de su vigente'))
    for cod, nom in (('Ítem', 'Nombre Ítem'), ('Asignación', 'Nombre Asignación')):
        x = base[[ 'Servicio', cod, nom]].dropna().drop_duplicates()
        dob = x.groupby(['Servicio', cod])[nom].nunique(); dob = dob[dob > 1]
        for (s, c), _ in dob.items():
            H.append(_h(G, f'Código con dos nombres ({cod})', AMARILLO, f'{s} {cod} {int(c)}: aparece con nombres distintos (' + ' / '.join(sorted(x[(x.Servicio == s) & (x[cod] == c)][nom])) + '); en Power BI se ven como dos líneas', s,
                        'Unifique el nombre en el catálogo o en la planilla'))
    con_monto = base[M.MONTOS].fillna(0).abs().sum(axis=1) > 0
    sin_reg = con_monto & (base['Código'].isna() | base['Región'].isna() | base['Mapa'].isna())
    if sin_reg.any():
        r = base[sin_reg]
        H.append(_h(G, 'Ejecución sin región', ROJO, f'{int(sin_reg.sum())} fila(s) con ejecución sin región o sin mapa (devengo {_m(r["Monto Devengo"].sum())} M$): no aparecen en Presupuesto Regional',
                    ', '.join(sorted(r.Servicio.unique())), 'Agregue la unidad en la hoja Regiones del catálogo (botón Agregar al catálogo)'))
    for c in ('Ley de Presupuestos', 'Presupuesto Vigente'):
        v = pd.to_numeric(bg[c], errors='coerce').dropna()
        if (v % 1 != 0).any():
            H.append(_h(G, f'Base Glosas: {c} con decimales', AMARILLO, f'Power BI lee "{c}" de glosas como número entero: {int((v % 1 != 0).sum())} valor(es) se redondean'))
    for (hoja, col), vals in LITERALES.items():
        df = base if hoja == 'Base Consolidada' else bg
        falta = [v for v in vals if v not in set(df[col].dropna().astype(str))]
        if falta:
            H.append(_h(G, 'Valores que filtra el informe Power BI', ROJO, f'{hoja}, columna "{col}": no aparece {", ".join(repr(v) for v in falta)}. Los visuales que filtran por ese valor quedarían vacíos',
                        '', 'Revise el nombre en el catálogo (debe ser idéntico al que usa el informe)'))
    if anterior is not None:
        try:
            pa = pd.read_excel(anterior, sheet_name='Base Consolidada'); pga = pd.read_excel(anterior, sheet_name='Base Glosas')
            pr_a = set(zip(pa.Servicio, pa.Programa)); pr_n = set(zip(base.Servicio, base.Programa))
            for s, p in sorted(pr_a - pr_n):
                H.append(_h(G, 'Programa que desaparece', ROJO, f'{s} {p}: estaba en el consolidado del mes anterior y no está en este', s, 'Revise la hoja Programas del catálogo y el informe del servicio'))
            for s, p in sorted(pr_n - pr_a):
                H.append(_h(G, 'Programa nuevo', AMARILLO, f'{s} {p}: no estaba en el consolidado del mes anterior', s))
            ga = set(zip(pga.Servicio, pga.Programa, pga.Glosa)); gn = set(zip(bg.Servicio, bg.Programa, bg.Glosa))
            for s, p, g in sorted(ga - gn):
                H.append(_h(G, 'Glosa que desaparece', AMARILLO, f'{s} {p} {g}: estaba en Base Glosas del mes anterior y no en esta', s))
            va = pa[pa['Presupuesto Vigente'].fillna(0) != 0]; vn = base[base['Presupuesto Vigente'].fillna(0) != 0]
            na = set(map(tuple, va[['Servicio', 'Nombre Asignación']].dropna().values)) if 'Nombre Asignación' in va else set()
            nn = set(map(tuple, base[['Servicio', 'Nombre Asignación']].dropna().values))   # en cualquier fila del mes (con o sin vigente)
            cambios = sorted(na - nn)
            if cambios:
                H.append(_h(G, 'Asignaciones con vigente que cambian de nombre o desaparecen', AMARILLO, f'{len(cambios)} asignación(es) con vigente el mes anterior no aparecen con el mismo nombre: '
                            + '; '.join(f'{s} {n}' for s, n in cambios[:8]) + ('…' if len(cambios) > 8 else ''), '', 'Si cambió el nombre en el catálogo o la planilla, las series mensuales se cortan'))
        except Exception as e:
            H.append(_h(G, 'Comparación con el mes anterior', AMARILLO, f'No se pudo leer el consolidado del mes anterior ({type(e).__name__})'))
    if not H: H.append(_h(G, 'Estructura para Power BI', VERDE, 'Columnas, tipos, nombres, regiones y valores de filtro del informe Power BI conformes'))
    return H


# ------------------------------------------------------------------ 6. huella del corte
def _sha(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for b in iter(lambda: f.read(1 << 20), b''): h.update(b)
    return h.hexdigest()


def huella_datos(base, bg):
    h = hashlib.sha256()
    for df in (base, bg):
        h.update(df.to_csv(index=False, float_format='%.6f').encode('utf-8'))
    return h.hexdigest()


def huella(anio, mes, base, bg, insumos, corte, version):
    """Con los mismos insumos, el mismo corte y la misma versión, el consolidado debe ser idéntico."""
    G = 'Huella del corte'; H = []
    ent = hashlib.sha256('|'.join(sorted(_sha(p) for p in insumos if p and os.path.exists(p))).encode() + f'|{corte:%Y-%m-%d}'.encode()).hexdigest()
    sal = huella_datos(base, bg)
    try:
        R = json.load(open(HUELLAS, encoding='utf-8')) if os.path.exists(HUELLAS) else {}
    except Exception:
        R = {}
    k = f'{anio}-{mes:02d}'; prev = [x for x in R.get(k, []) if x['entradas'] == ent and x['version'] == version]
    if prev and prev[-1]['salida'] != sal:
        H.append(_h(G, 'Reproceso no idéntico', ROJO, f'Con los mismos insumos, corte y versión, la generación del {prev[-1]["fecha"]} produjo un consolidado distinto',
                    '', 'Avise a soporte del agente: el resultado debe ser siempre el mismo'))
    elif prev:
        H.append(_h(G, 'Reproceso idéntico', VERDE, f'Mismo resultado que la generación del {prev[-1]["fecha"]} (mismos insumos, corte y versión)'))
    else:
        H.append(_h(G, 'Huella registrada', VERDE, f'Huella del consolidado: {sal[:16]}… (insumos {ent[:16]}…). Si se vuelve a generar con los mismos insumos, el resultado debe ser idéntico'))
    R.setdefault(k, []).append({'fecha': datetime.datetime.now().strftime('%d-%m-%Y %H:%M'), 'entradas': ent, 'salida': sal, 'version': version, 'corte': f'{corte:%d-%m-%Y}'})
    R[k] = R[k][-20:]
    try:
        os.makedirs(os.path.dirname(HUELLAS), exist_ok=True); json.dump(R, open(HUELLAS, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    except Exception:
        pass
    return H, ent, sal


# ------------------------------------------------------------------ conjunto
def validar_corte(K, base, bg, ideas, pvs, corte, anio, mes, insumos, version, pares, anterior=None):
    traza_pv = pd.concat(M.TRAZA_PV, ignore_index=True) if M.TRAZA_PV else None
    H = []
    H += contrato_entrada(M.BRUTOS)
    H += contrato_planillas(pvs, traza_pv)
    H += cuadratura(base, M.CUADRATURA, traza_pv)
    H += sombra(K, M.BRUTOS, base)
    H += coherencia(base, traza_pv, pares)
    H += contrato_salida(base, bg, ideas, K, anterior)
    hh, ent, sal = huella(anio, mes, base, bg, insumos, corte, version); H += hh
    df = pd.DataFrame(H, columns=['Grupo', 'Control', 'Criticidad', 'Servicio', 'Detalle', 'Qué revisar'])
    return {'hallazgos': df, 'traza_pv': traza_pv, 'traza': M.TRAZA.get('ejecucion'), 'huella': sal, 'huella_insumos': ent}


def semaforo(df):
    """Una fila por grupo: el peor estado y el número de hallazgos."""
    orden = {ROJO: 0, AMARILLO: 1, VERDE: 2}; out = []
    for g in GRUPOS:
        x = df[df.Grupo == g]
        if not len(x): continue
        peor = min(x.Criticidad, key=lambda c: orden[c])
        n_r, n_a = int((x.Criticidad == ROJO).sum()), int((x.Criticidad == AMARILLO).sum())
        res = 'Sin observaciones' if peor == VERDE else ', '.join(t for t in (f'{n_r} rojo(s)' if n_r else '', f'{n_a} amarillo(s)' if n_a else '') if t)
        out.append({'Control': g, 'Estado': peor, 'Resultado': res})
    return pd.DataFrame(out)


def ruta_validacion(carpeta_mes, anio, mes):
    return os.path.join(carpeta_mes, 'MINAGRI', f'MINAGRI Validación a {M_MESES[mes - 1]} {anio}.xlsx')


M_MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
COLOR = {ROJO: 'F8D7DA', AMARILLO: 'FFF3CD', VERDE: 'D1E7DD'}


def guardar(V, carpeta_mes, anio, mes, corte, version, consolidado, pbix=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    ruta = ruta_validacion(carpeta_mes, anio, mes)
    wb = Workbook(write_only=False); ws = wb.active; ws.title = 'Semáforo'
    azul = '1F4E5F'; neg = Font(bold=True)
    ws['A1'] = 'MINISTERIO DE AGRICULTURA – UPROGE'; ws['A1'].font = Font(bold=True, size=13, color=azul)
    ws['A2'] = f'Validación del corte · Consolidado a {M_MESES[mes - 1]} {anio}'; ws['A2'].font = Font(bold=True, size=12)
    info = [('Consolidado', os.path.basename(consolidado)), ('Fecha de corte de decretos', f'{corte:%d-%m-%Y}'),
            ('Generado', f'{datetime.datetime.now():%d-%m-%Y %H:%M} · agente versión {version}'),
            ('Huella del consolidado', V['huella']), ('Huella de los insumos', V['huella_insumos']),
            ('Cómo leer', 'Verde: sin observaciones. Amarillo: revisar. Rojo: cambia cifras o rompe el informe; corríjalo o explíquelo en la pestaña Revisión antes de publicar.')]
    f = 4
    for k, v in info: ws.cell(f, 1, k).font = neg; ws.cell(f, 2, v); f += 1
    f += 1
    sem = semaforo(V['hallazgos'])
    for j, h in enumerate(['Control', 'Estado', 'Resultado'], 1):
        c = ws.cell(f, j, h); c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor=azul)
    for _, r in sem.iterrows():
        f += 1; ws.cell(f, 1, r.Control); c = ws.cell(f, 2, r.Estado); c.fill = PatternFill('solid', fgColor=COLOR[r.Estado]); ws.cell(f, 3, r.Resultado)
    if pbix:
        f += 2; ws.cell(f, 1, 'Revisión del informe Power BI').font = Font(bold=True, color=azul)
        for x in pbix:
            f += 1; c = ws.cell(f, 1, x['Criticidad']); c.fill = PatternFill('solid', fgColor=COLOR[x['Criticidad']]); ws.cell(f, 2, x['Elemento']); ws.cell(f, 3, x['Detalle'])
    for col, w in zip('ABC', (34, 70, 90)): ws.column_dimensions[col].width = w
    # hallazgos
    wh = wb.create_sheet('Hallazgos'); hz = V['hallazgos']
    wh.append(list(hz.columns))
    for c in wh[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor=azul)
    for r in hz.itertuples(index=False):
        wh.append(list(r)); fila = wh.max_row
        wh.cell(fila, 3).fill = PatternFill('solid', fgColor=COLOR[r.Criticidad])
        for j in (5, 6): wh.cell(fila, j).alignment = Alignment(wrap_text=True, vertical='top')
    for col, w in zip('ABCDEF', (26, 40, 11, 18, 100, 60)): wh.column_dimensions[col].width = w
    wh.freeze_panes = 'A2'
    # trazabilidad
    for nombre, df in (('Trazabilidad ejecución', V.get('traza')), ('Trazabilidad presupuesto', V.get('traza_pv'))):
        if df is None or not len(df): continue
        d = df.copy()
        if nombre == 'Trazabilidad presupuesto':
            d = d.rename(columns={'Fecha Documento': 'Fecha aplicada', 'Fecha Planilla': 'Fecha en la planilla'})
        wt = wb.create_sheet(nombre); wt.append(list(d.columns))
        for c in wt[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor=azul)
        for r in d.itertuples(index=False):
            wt.append([None if (isinstance(v, float) and pd.isna(v)) or v is pd.NA or v is pd.NaT else (v.to_pydatetime() if isinstance(v, pd.Timestamp) else v) for v in r])
        wt.freeze_panes = 'A2'; wt.auto_filter.ref = wt.dimensions
        for j, c in enumerate(d.columns, 1):
            wt.column_dimensions[wt.cell(1, j).column_letter].width = max(12, min(48, len(str(c)) + 4))
            if 'Fecha' in c:
                for cell in wt.iter_cols(min_col=j, max_col=j, min_row=2):
                    for x in cell: x.number_format = 'DD-MM-YYYY'
    if 'Trazabilidad ejecución' in wb.sheetnames:
        wn = wb.create_sheet('Cómo usar la trazabilidad', 1)
        for t in ('Trazabilidad ejecución: una fila por fila de la hoja Base Consolidada (columna "Fila Base Consolidada").',
                  'Las filas de ejecución indican el informe bruto y la fila de Excel de donde salieron; las de presupuesto vienen de la planilla de Presupuesto Vigente.',
                  'Trazabilidad presupuesto: cada fila de la hoja "Base Presupuesto" de cada planilla, con la fecha aplicada al decreto y si quedó dentro del vigente del mes.',
                  'Para explicar una cifra: filtre por servicio, programa y código, y sume las filas de origen.'):
            wn.append([t])
        wn.column_dimensions['A'].width = 140
    try:
        wb.save(ruta)
    except PermissionError:
        raise PermissionError(f'No pude guardar {os.path.basename(ruta)}. ¿Está abierto en Excel? Ciérrelo y vuelva a generar.')
    return ruta
