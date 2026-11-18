"""Motor del Agente de Ejecución Presupuestaria MINAGRI.
Convierte los informes brutos (SIGFE II, SIGFIN, ERP CONAF) + las planillas de Presupuesto Vigente
en las hojas Base Consolidada, Base Glosas e Ideas Fuerza que lee el informe Power BI.
No modifica ningún archivo de entrada."""
import os, re, glob, openpyxl, pandas as pd

MONTOS=['Monto Requerimiento','Monto Compromiso','Monto Devengo']
PPTO=['Presupuesto Ley','Presupuesto Vigente']
SERV=['01 SUBSE','02 ODEPA','03 INDAP','04 SAG','05 CONAF','06 CNR']
COLS=['Servicio','Código','Región','Mapa','Programa','Tipo','Subtítulo','Nombre Subtítulo','Ítem','Nombre Ítem','Asignación',
      'Nombre Asignación','Subasignación','Nombre Subasignación']+PPTO+MONTOS
COLS_GLOSAS=['Servicio','Programa','Código Glosa','Glosa','Ley de Presupuestos','Presupuesto Vigente']+MONTOS
HOJAS_AUX={'factores','base limpia $','base limpia m$','revisión','revisión base','revision','revision base','verificación',
           'verificación base','tabla emergencias','hoja1','clasificador ingresos'}
INSTITUCIONES=[(0,'MINAGRI'),(1,'SUBSE'),(2,'ODEPA'),(3,'INDAP'),(4,'SAG'),(5,'CONAF'),(6,'CNR')]

class ErrorAgente(Exception): pass

def norm(x):
    if x is None or (isinstance(x,float) and pd.isna(x)): return None
    if isinstance(x,float) and x.is_integer(): x=int(x)
    s=str(x).strip(); return s or None

# ------------------------------------------------------------------ catálogo
class Catalogo:
    def __init__(self,path):
        if not os.path.exists(path): raise ErrorAgente(f'No encuentro el Catálogo Maestro en: {path}')
        x=pd.read_excel(path,sheet_name=None,dtype=str)
        for h in ['Servicios','Regiones','Programas','Clasificador','Clasificador por servicio','Ingresos SIGFIN','Glosas','Glosas SAG por U.E.','Mapa','Corte decretos']:
            if h not in x: raise ErrorAgente(f'Al Catálogo Maestro le falta la hoja "{h}"')
        rg=x['Regiones'].dropna(subset=['Valor en el informe bruto'])
        malas=[i+2 for i,v in zip(rg.index,rg['Código región']) if pd.isna(pd.to_numeric(v,errors='coerce'))]
        if malas: raise ErrorAgente(f'Catálogo Maestro, hoja "Regiones": la columna "Código región" debe ser un número (filas {", ".join(map(str,malas[:10]))}). Corríjalo y vuelva a generar')
        self.reg={(r.Servicio,r['Valor en el informe bruto']):(int(float(r['Código región'])),r.Región) for _,r in rg.iterrows()}
        p=x['Programas'].fillna('')
        self.prog={(r.Servicio,r['Valor en el informe bruto']):(r.Tratamiento.strip(),r.Programa.strip()) for _,r in p.iterrows() if r['Valor en el informe bruto']}
        self.nom={(r.Nivel,r.Código):r.Nombre for _,r in x['Clasificador'].dropna(subset=['Código']).iterrows()}
        self.nom_s={(r.Servicio,r.Nivel,r.Código):r.Nombre for _,r in x['Clasificador por servicio'].dropna(subset=['Código']).iterrows()}
        self.ing={(r.Servicio,r['Catálogo en el informe bruto']):r['Código clasificador'] for _,r in x['Ingresos SIGFIN'].dropna(subset=['Catálogo en el informe bruto']).iterrows()}
        self.glo={(r.Servicio,r['Código clasificador']):(r.Glosa,int(float(r['Código Glosa']))) for _,r in x['Glosas'].dropna(subset=['Código clasificador']).iterrows()}
        self.glo_ue={(r.Servicio,r['U.E. (SIGFIN)']):(r.Glosa,int(float(r['Código Glosa']))) for _,r in x['Glosas SAG por U.E.'].dropna(subset=['U.E. (SIGFIN)']).iterrows()}
        self.mapa={int(float(r['Código región'])):r['Mapa (Power BI)'] for _,r in x['Mapa'].iterrows()}
        c=pd.read_excel(path,sheet_name='Corte decretos',dtype={'Mes':str}).dropna(subset=['Mes'])
        # fecha escrita como texto (DD-MM-AAAA) o como fecha de Excel
        self.corte={str(m).strip()[:7]:(pd.to_datetime(str(f).strip(),dayfirst=True) if isinstance(f,str) else pd.Timestamp(f)) for m,f in zip(c.Mes,c['Fecha de corte de decretos']) if pd.notna(f) and str(f).strip()}
        self.llave=dict(zip(x['Servicios'].Servicio,x['Servicios']['Llave para la región']))
        sv=x['Servicios'].fillna('')
        self.defecto={(r.Servicio,t):(str(r.get(f'Programas no registrados ({t})','') or 'Excluir')).strip() for _,r in sv.iterrows() for t in ('Gastos','Ingresos')}
        self.p01={s:next((p for (ss,_),(t,p) in self.prog.items() if ss==s and p.startswith('P01')),'') for s in self.llave}
    def nombre(self,s,nivel,cod): return self.nom_s.get((s,nivel,cod), self.nom.get((nivel,cod)))

# ------------------------------------------------------------------ lectura de brutos (.xlsx / .xls)
MARCAS=('Código Unidad Ejecutora','A.U.')
def _filas(path):
    """Devuelve las filas de la hoja que contiene el informe bruto (la primera hoja cuyo encabezado sea de SIGFE o SIGFIN/ERP).
    Las demás hojas (Base Limpia, Factores, Revisión, etc.) se ignoran."""
    if path.lower().endswith('.xls'):
        import xlrd
        wb=xlrd.open_workbook(path)
        for sh in wb.sheets():
            if any(sh.nrows>i and str(sh.cell_value(i,0)).strip() in MARCAS for i in range(min(60,sh.nrows))):
                return [sh.row_values(i) for i in range(sh.nrows)]
        return []
    wb=openpyxl.load_workbook(path,read_only=True,data_only=True)
    for n in wb.sheetnames:
        cab=list(wb[n].iter_rows(max_row=60,max_col=1,values_only=True))
        if any(r and str(r[0]).strip() in MARCAS for r in cab): return list(wb[n].iter_rows(values_only=True))
    return []

def leer_bruto(path):
    rows=_filas(path)
    h=next((i for i,r in enumerate(rows[:60]) if r and str(r[0]).strip() in ('Código Unidad Ejecutora','A.U.')),None)
    if h is None: raise ErrorAgente(f'{os.path.basename(path)}: no reconozco el formato (no encuentro el encabezado "Código Unidad Ejecutora" de SIGFE ni "A.U." de SIGFIN/ERP en las primeras 60 filas).')
    cols=[];seen={}
    for c in rows[h]:
        c=str(c).strip() if c not in (None,'') else 'None'; seen[c]=seen.get(c,0)+1; cols.append(c if seen[c]==1 else f'{c}.{seen[c]-1}')
    df=pd.DataFrame([list(r)+[None]*(len(cols)-len(r)) for r in rows[h+1:]],columns=cols)
    df.attrs['fila0']=h+2          # fila de Excel (1 = primera) de la primera fila de datos: para la trazabilidad
    return df, ('SIGFIN' if cols[0]=='A.U.' else 'SIGFE')

def por_prefijo(df,pref,archivo):
    for c in df.columns:
        v=df[c].dropna().astype(str)
        if len(v) and v.str.startswith(pref).mean()>0.9: return df[c]
    raise ErrorAgente(f'{archivo}: no encuentro la columna de catálogo "{pref}…"')

def _col(df,nombre,archivo):
    if nombre not in df.columns: raise ErrorAgente(f'{archivo}: falta la columna "{nombre}" en el informe bruto')
    return df[nombre]

def ejecucion(K,path,s,t,av):
    a=os.path.basename(path); df,fmt=leer_bruto(path)
    BRUTOS[(s,t)]={'df':df,'formato':fmt,'archivo':a,'ruta':path}   # para el contrato de entrada y el motor sombra
    fila0=df.attrs.get('fila0',2)
    if fmt=='SIGFIN':
        df=df[pd.to_numeric(_col(df,'A.U.',a),errors='coerce').fillna(0)!=0]
        out=pd.DataFrame({'Servicio':s,'Tipo':t},index=df.index)
        llave=df.iloc[:,list(df.columns).index('C.F')+1].map(norm)
        cf=_col(df,'C.F',a).map(norm)   # código del centro financiero: si está en el catálogo, manda sobre la descripción
        prog=_col(df,'P.P.',a).map(norm); cat=_col(df,'Catalogo',a).map(norm)
        if t=='Ingresos':
            sin=sorted(set(cat.dropna())-{k[1] for k in K.ing if k[0]==s})
            for v in sin: _pend('Ingreso',s,t,v)
            if sin: av.append(('Revisar',f'{s} Ingresos: catálogos sin equivalencia en la hoja "Ingresos SIGFIN" del catálogo (quedan sin clasificador): {", ".join(sin)}'))
            cat=cat.map(lambda v: K.ing.get((s,v)))
        num=lambda c: pd.to_numeric(_col(df,c,a),errors='coerce').fillna(0)/1000
        out['Monto Requerimiento']=num('Presupuesto'); out['Monto Compromiso']=num('Compromiso' if t=='Gastos' else 'Devengo'); out['Monto Devengo']=num('Devengo')
        out['_ue']=_col(df,'U.E.',a).map(norm)
    else:
        df=df[_col(df,'Código Concepto',a).notna()]
        out=pd.DataFrame({'Servicio':s,'Tipo':t},index=df.index)
        # La región se busca con la columna que mejor calza con la hoja Regiones del catálogo
        # (SUBSE/INDAP/CONAF: descripción de la unidad ejecutora; ODEPA: código; CNR: unidad demandante).
        cands=[_col(df,'Descripción Unidad Ejecutora',a),_col(df,'Código Unidad Ejecutora',a)]
        try: cands.append(por_prefijo(df,'UnidadesDemandantes',a))
        except ErrorAgente: pass
        cands=[c.map(norm) for c in cands]
        llave=max(cands,key=lambda c: sum((s,v) in K.reg for v in c)); prog=por_prefijo(df,'ProgramaPresupuestario',a).map(norm); cat=df['Código Concepto'].map(norm)
        for m in MONTOS: out[m]=pd.to_numeric(_col(df,m,a),errors='coerce').fillna(0)/1000
        out['_ue']=None
        cf=None
    out['_archivo']=a; out['_fila']=out.index+fila0
    rg=[K.reg.get((s,v)) for v in llave]
    if cf is not None and s=='05 CONAF' and not any(k[0]==s and str(k[1]).isdigit() for k in K.reg):
        av.append(('Revisar','05 CONAF: el Catálogo Maestro no trae las regiones por código C.F. (versión 1.4). Instale el catálogo actualizado: '
                             'con el catálogo anterior la distribución regional de CONAF queda errada en 12 oficinas'))
    if cf is not None and any((s,c) in K.reg for c in cf.dropna().unique()):
        rc=[K.reg.get((s,c)) for c in cf]
        dist=sorted({f'C.F. {c} {v}: catálogo por código {x[0]}, por nombre {y[0]}' for c,v,x,y in zip(cf,llave,rc,rg) if x and y and x[0]!=y[0]})
        if dist: av.append(('Revisar',f'{s} {t}: la región por código C.F. no coincide con la región por nombre de oficina en la hoja "Regiones" (se usó la del código): {"; ".join(dist)}'))
        nuevos_cf=sorted({f'{c} {v}' for c,v,x in zip(cf,llave,rc) if c and not x})
        for c,v in sorted({(c,v) for c,v,x in zip(cf,llave,rc) if c and not x}): _pend('Región CF',s,t,c,nota=v)
        if nuevos_cf: av.append(('Revisar',f'{s} {t}: centros financieros sin código C.F. en la hoja "Regiones" (se usó el nombre de oficina): {"; ".join(nuevos_cf)}'))
        rg=[x or y for x,y in zip(rc,rg)]
        if s=='05 CONAF' and any(c=='1701' and x and x[0]!=17 for c,x in zip(cf,rg)):
            av.append(('Revisar','05 CONAF: Oficina Central (C.F. 1701) quedó fuera de Nivel Central (17). Revise la hoja "Regiones" del catálogo: la tabla puede haberse corrido'))
    out['Código']=[r[0] if r else None for r in rg]; out['Región']=[r[1] if r else None for r in rg]
    sinreg=sorted({v for v,r in zip(llave,rg) if r is None and v})
    for v in sinreg: _pend('Región',s,t,v)
    if sinreg: av.append(('Revisar',f'{s} {t}: unidades sin región en la hoja "Regiones" del catálogo (quedan sin región): {"; ".join(sinreg)}'))
    tr=[K.prog.get((s,v),('NUEVO','')) for v in prog]
    nuevos=sorted({v for v,(x,_) in zip(prog,tr) if x=='NUEVO' and v})
    for v in nuevos: _pend('Programa',s,t,v)
    if nuevos:
        ok=K.defecto.get((s,t),'Excluir')=='Sumar a P01' and bool(K.p01.get(s))
        tr=[((('Sumar a P01',K.p01[s]) if ok else ('Excluir','')) if x=='NUEVO' else (x,p)) for x,p in tr]
        av.append(('Revisar',f'{s} {t}: programas que no están en la hoja "Programas" del catálogo; se aplicó el tratamiento por defecto ({"sumados a P01" if ok else "excluidos"}). Agréguelos al catálogo: {"; ".join(nuevos)}'))
    out['Programa']=[p if x in ('Incluir','Sumar a P01') and p else None for x,p in tr]
    exc=out[out.Programa.isna()]
    CUADRATURA[(s,t)]={'Devengo informe':float(out['Monto Devengo'].sum()),'Devengo excluido por programa':float(exc['Monto Devengo'].sum()),
                       **{f'{m} informe':float(out[m].sum()) for m in MONTOS},**{f'{m} excluido':float(exc[m].sum()) for m in MONTOS}}
    if len(exc): av.append(('Info',f'{s} {t}: {len(exc)} filas excluidas por programa (Devengo {exc["Monto Devengo"].sum():,.0f}; Requerimiento {exc["Monto Requerimiento"].sum():,.0f} miles de $)'.replace(',','.')))
    out=out[out.Programa.notna()].copy(); cat=cat.loc[out.index]
    out['Subtítulo']=cat.str[:2]; out['Ítem']=cat.str[:4]; out['Asignación']=cat.str[:7]; out['Subasignación']=cat.str[:10]
    for n,c in [('Subtítulo','Nombre Subtítulo'),('Ítem','Nombre Ítem'),('Asignación','Nombre Asignación'),('Subasignación','Nombre Subasignación')]:
        out[c]=[K.nombre(s,n,v) for v in out[n]]
    # asignación "000" (ítem que el servicio no abre en asignaciones): la ejecución queda a nivel de ítem, igual que el
    # presupuesto vigente y que el consolidado oficial: se conservan los códigos y la asignación/subasignación van sin
    # nombre. Así Power BI muestra una sola línea por ítem (con vigente y ejecución) y no dos líneas separadas.
    cero=out['Asignación'].astype(str).str.fullmatch(r'\d{4}000')
    if cero.any(): out.loc[cero,['Nombre Asignación','Nombre Subasignación']]=None
    for n in ('Subtítulo','Ítem'):
        f=sorted(out.loc[out['Nombre '+n].isna(),n].dropna().unique())
        for c in f: _pend('Nombre',s,t,c,nivel=n,sugerido='')
        if f: av.append(('Revisar',f'{s} {t}: códigos de {n} sin nombre en la hoja "Clasificador": {", ".join(f)}'))
    return out

def glosas_ejecucion(K,d):
    if d.empty: return d.assign(Glosa=None,**{'Código Glosa':None})   # todo el informe quedó excluido por programa
    s=d.Servicio.iat[0]
    g=[K.glo.get((s,sa)) or K.glo.get((s,a)) for sa,a in zip(d['Subasignación'],d['Asignación'])]
    if s=='04 SAG': g=[K.glo_ue.get((s,u)) or x for x,u in zip(g,d['_ue'])]   # U.E. 1001/1002/2001/2002: toda la U.E. va a su glosa, incluida su deuda flotante (criterio de agosto y octubre)
    d=d.assign(Glosa=[x[0] if x else None for x in g],**{'Código Glosa':[x[1] if x else None for x in g]})
    return d[d.Glosa.notna()]

# ------------------------------------------------------------------ presupuesto vigente
def planillas_pv(carpeta):
    pv={}
    for s in SERV:
        c=[f for f in glob.glob(os.path.join(carpeta,'*.xlsx')) if os.path.basename(f).startswith(s[:2]) and 'vigente' in f.lower() and not os.path.basename(f).startswith('~$')]
        if len(c)!=1: raise ErrorAgente(f'En "{carpeta}" debe haber exactamente una planilla "{s} Presupuesto Vigente"; hay {len(c)}')
        pv[s]=c[0]
    return pv

FECHAS_DECRETO={}   # {'1289': Timestamp} fechas de decretos a aplicar (las fija proceso.py tras verificar en DIPRES)
NIVELES_COD=(('Subtítulo','Nombre Subtítulo',2),('Ítem','Nombre Ítem',4),('Asignación','Nombre Asignación',7),('Subasignación','Nombre Subasignación',10))

def base_presupuesto(path):
    """Hoja Base Presupuesto. Si una columna de código viene vacía (fórmula sin calcular, p. ej. filas agregadas
    sin recalcular el libro), el código se toma del número con que empieza el nombre ("24 TRANSFERENCIAS…" -> 24).
    La columna _completado marca esas filas para avisar."""
    b=pd.read_excel(path,sheet_name='Base Presupuesto',dtype={'Subtítulo':str,'Ítem':str,'Asignación':str,'Subasignación':str})
    b['_completado']=False
    if FECHAS_DECRETO and 'Numero Documento' in b and 'Fecha Documento' in b:
        # fecha a aplicar según el criterio (toma de razón en DIPRES); se conserva la de la planilla. Solo en memoria.
        b['Fecha Planilla']=b['Fecha Documento']
        n=pd.to_numeric(b['Numero Documento'],errors='coerce').map(lambda x:str(int(x)) if pd.notna(x) else '')
        f=n.map(FECHAS_DECRETO); ok=(b.Documento=='Decreto')&f.notna()
        b.loc[ok,'Fecha Documento']=pd.to_datetime(f[ok])
    for cod,nom,largo in NIVELES_COD:
        if cod not in b or nom not in b: continue
        vacio=b[cod].isna()|(b[cod].astype(str).str.strip().isin(['','nan','None']))
        desde=b[nom].astype(str).str.extract(r'^\s*(\d+)\s')[0]
        ok=vacio&desde.notna()&(desde.str.len()==largo)
        b.loc[ok,cod]=desde[ok]; b.loc[ok,'_completado']=True
    return b

def registro_decretos(pvs):
    out=[]
    for p in pvs.values():
        l=base_presupuesto(p); out.append(l[(l.Documento=='Decreto')&l['Numero Documento'].notna()][['Numero Documento','Fecha Documento']])
    r=pd.concat(out).groupby('Numero Documento')['Fecha Documento'].min(); r.index=r.index.map(lambda x:str(int(x))); return r

def presupuesto(path,s,corte,av):
    p=base_presupuesto(path); p=p[p['Monto M$'].notna()]
    c=p[p['_completado']]
    if len(c):
        docs=', '.join(f'{d} N°{int(n)}' if pd.notna(n) else str(d) for (d,n),_ in c.groupby(['Documento','Numero Documento'],dropna=False))
        av.append(('Info',f'{s}: {len(c)} fila(s) de la planilla de Presupuesto Vigente tenían vacías las columnas de código '
                          f'(fórmulas sin calcular: {docs}). El agente tomó el código desde el nombre. '
                          'Para dejar la planilla al día: ábrala en Excel, presione Ctrl+Alt+F9 y guarde'))
    inc=(p.Documento!='Decreto')|(p['Fecha Documento']<=corte)
    t=p.assign(**{'Servicio':s,'Fila planilla':p.index+2,'Incluido en el vigente':inc.map({True:'Sí',False:'No (posterior al corte)'})})
    TRAZA_PV.append(t[[c for c in ['Servicio','Fila planilla','Documento','Numero Documento','Fecha Planilla','Fecha Documento','Incluido en el vigente','Programa','Tipo',
                                   'Subtítulo','Ítem','Asignación','Subasignación','Nombre Asignación','Monto M$'] if c in t.columns]])
    fuera=p[(p.Documento=='Decreto')&(p['Fecha Documento']>corte)]
    if len(fuera): av.append(('Info',f'{s}: decretos posteriores al corte, no incluidos: '+', '.join(f'N°{int(n)} ({f:%d-%m-%Y})' for n,f in fuera.groupby('Numero Documento')['Fecha Documento'].min().items())))
    p=p[(p.Documento!='Decreto')|(p['Fecha Documento']<=corte)]
    p=p.assign(**{'Presupuesto Ley':p['Monto M$'].where(p.Documento!='Decreto',0),'Presupuesto Vigente':p['Monto M$']})
    k=['Programa','Tipo','Subtítulo','Nombre Subtítulo','Ítem','Nombre Ítem','Asignación','Nombre Asignación','Subasignación','Nombre Subasignación']
    g=p.groupby(k,dropna=False)[PPTO].sum().reset_index()
    g.insert(0,'Servicio',s); g['Código']=17; g['Región']='Nivel Central'; return g

def glosas_presupuesto(path,s,corte,R,av):
    wb=openpyxl.load_workbook(path); wv=openpyxl.load_workbook(path,data_only=True)
    if 'Base Glosas (No Modificar)' not in wb.sheetnames or 'Glosas' not in wb.sheetnames:
        raise ErrorAgente(f'{os.path.basename(path)}: faltan las hojas "Glosas" o "Base Glosas (No Modificar)"')
    bg=wb['Base Glosas (No Modificar)']; gs=wv['Glosas']; bgv=wv['Base Glosas (No Modificar)']; filas=[]
    for r in range(2,bg.max_row+1):
        f=bg.cell(r,5).value
        if not isinstance(f,str) or 'Glosas!' not in f: continue
        row=int(re.search(r'Glosas!\$?D\$?(\d+)',f).group(1)); h=row
        while h>1 and not str(gs.cell(h,4).value or '').startswith('Ley de Presupuestos'): h-=1
        ley=gs.cell(row,4).value or 0; vig=ley
        for c in range(5,17):
            lab=str(gs.cell(h+1,c).value or ''); v=gs.cell(row,c).value or 0
            if not v: continue
            mm=re.search(r'(\d{2})-(\d{2})-(\d{4})',lab); num=re.search(r'N°\s*(\d+)',lab)
            if num and num.group(1) in FECHAS_DECRETO: fe=FECHAS_DECRETO[num.group(1)]
            elif mm: fe=pd.Timestamp(f'{mm.group(3)}-{mm.group(2)}-{mm.group(1)}')
            elif num and num.group(1) in R.index: fe=R[num.group(1)]
            else: av.append(('Revisar',f'{s}: el decreto "{lab.strip()}" de la hoja Glosas no tiene fecha (ni en el encabezado ni en Base Presupuesto); se incluyó en el vigente de glosas')); fe=corte
            if fe<=corte: vig+=v
        filas.append({'Servicio':s,'Programa':bgv.cell(r,2).value,'Código Glosa':bgv.cell(r,3).value,'Glosa':bgv.cell(r,4).value,'Ley de Presupuestos':ley,'Presupuesto Vigente':vig})
    return pd.DataFrame(filas)

def sugerir_corte(pvs,oficial):
    """Busca la fecha de corte que reproduce el Presupuesto Vigente del consolidado oficial (Gastos, por servicio)."""
    O=pd.read_excel(oficial,sheet_name='Base Consolidada'); obj=O[O.Tipo=='Gastos'].groupby('Servicio')['Presupuesto Vigente'].sum()
    B={s:base_presupuesto(p) for s,p in pvs.items()}
    fechas=sorted({f for b in B.values() for f in b.loc[b.Documento=='Decreto','Fecha Documento'].dropna()})
    cands=[fechas[0]-pd.Timedelta(days=1)]+fechas
    def err(c):
        n=0
        for s,b in B.items():
            b=b[(b.Tipo=='Gastos')&b['Monto M$'].notna()]; v=b[(b.Documento!='Decreto')|(b['Fecha Documento']<=c)]['Monto M$'].sum()
            n+=abs(v-obj.get(s,0))>1
        return n
    res=[(c,err(c)) for c in cands]; mejor=min(e for _,e in res)
    i=next(i for i,(c,e) in enumerate(res) if e==mejor)
    desde=res[i][0]; hasta=(res[i+1][0]-pd.Timedelta(days=1)) if i+1<len(res) else None
    return desde,hasta,mejor

# ------------------------------------------------------------------ orquestación
def _pend(tipo,s,t,valor,**kw):
    for d in PENDIENTES:   # la llave del catálogo no distingue Gastos de Ingresos: un solo pendiente
        if (d['tipo'],d['servicio'],d['valor'])==(tipo,s,valor):
            if t not in d['clase']: d['clase']=d['clase']+' e '+t
            return
    PENDIENTES.append(dict(tipo=tipo,servicio=s,clase=t,valor=valor,**kw))

def ubicar_archivos(carpeta):
    hall={}
    for f in sorted(glob.glob(os.path.join(carpeta,'*.xls'))+glob.glob(os.path.join(carpeta,'*.xlsx'))):
        b=os.path.basename(f)
        if b.startswith('~$'): continue
        for s in SERV:
            if b.startswith(s[:2]) and s.split()[1].lower() in b.lower():
                for t in ('Gastos','Ingresos'):
                    if t.lower() in b.lower(): hall.setdefault((s,t),[]).append(f)
    faltan=[f'{s} {t}' for s in SERV for t in ('Gastos','Ingresos') if (s,t) not in hall]
    dobles=[f'{s} {t}: '+' / '.join(os.path.basename(x) for x in v) for (s,t),v in hall.items() if len(v)>1]
    if faltan or dobles:
        msg=[]
        if faltan: msg.append('Faltan informes: '+', '.join(faltan))
        if dobles: msg.append('Hay más de un archivo para: '+'; '.join(dobles))
        raise ErrorAgente(' | '.join(msg)+f'\nCarpeta revisada: {carpeta}\nCada archivo debe empezar con el número del servicio y decir Gastos o Ingresos, p.ej. "05 CONAF Gastos a Octubre.xlsx".')
    return {k:v[0] for k,v in hall.items()}

BRUTOS={}       # informes brutos leídos, por (servicio, tipo): para el contrato de entrada y el motor sombra
TRAZA_PV=[]     # filas de las planillas de Presupuesto Vigente, incluidas o no en el vigente del mes
TRAZA={}        # trazabilidad de la última generación: {'ejecucion': DataFrame}
PENDIENTES=[]   # lo que falta en el Catálogo Maestro (programas, regiones, nombres, ingresos), para agregarlo desde la ventana
CUADRATURA={}   # devengo leído de cada informe bruto, para el control de cuadratura de la revisión

def generar(K,carpeta_mes,pvs,corte,ideas_df,log=print):
    av=[]; E=[]; GE=[]; P=[]; GP=[]; ctrl=[]; CUADRATURA.clear(); PENDIENTES.clear(); BRUTOS.clear(); TRAZA_PV.clear(); TRAZA.clear()
    arch=ubicar_archivos(carpeta_mes); R=registro_decretos(pvs)
    for s in SERV:
        for t in ('Gastos','Ingresos'):
            log(f'  Leyendo {os.path.basename(arch[(s,t)])} ...')
            d=ejecucion(K,arch[(s,t)],s,t,av); E.append(d)
            if t=='Gastos': GE.append(glosas_ejecucion(K,d))
        log(f'  Presupuesto vigente y glosas {s} ...')
        P.append(presupuesto(pvs[s],s,corte,av)); GP.append(glosas_presupuesto(pvs[s],s,corte,R,av))
    base=pd.concat(P+E,ignore_index=True); base['Mapa']=base['Código'].map(K.mapa)
    tz=base[['Servicio','Tipo','Programa']].copy(); tz.insert(0,'Fila Base Consolidada',range(2,len(base)+2))
    tz['Origen']=base['_archivo'].where(base['_archivo'].notna(),'Planilla de Presupuesto Vigente (hoja Base Presupuesto)')
    tz['Fila de origen']=base['_fila'].astype('Int64')
    tz['Código clasificador']=base['Subasignación'].fillna(base['Asignación']).fillna(base['Ítem']).fillna(base['Subtítulo'])
    for m in PPTO+MONTOS: tz[m]=base[m]
    TRAZA['ejecucion']=tz
    base=base[COLS]
    for c in ('Subtítulo','Ítem','Asignación','Subasignación'): base[c]=pd.to_numeric(base[c],errors='coerce').astype('Int64')
    bg=pd.concat(GP+[g[['Servicio','Programa','Código Glosa','Glosa']+MONTOS] for g in GE],ignore_index=True)[COLS_GLOSAS]
    ctrl=base.groupby(['Servicio','Tipo'])[PPTO+MONTOS].sum().reset_index()
    ctrl['% Avance (Devengo/Vigente)']=(ctrl['Monto Devengo']/ctrl['Presupuesto Vigente']).round(4)
    return base,bg,ideas_df,ctrl,av

def plantilla_ideas():
    return pd.DataFrame([(c,n,'[Pendiente de redacción]') for c,n in INSTITUCIONES],columns=['Código Institución','Institución','Ideas Fuerza'])

def validar(base,bg,oficial):
    """Compara contra un consolidado oficial ya existente (solo para pruebas)."""
    D={'Subtítulo':str,'Ítem':str,'Asignación':str,'Subasignación':str}
    O=pd.read_excel(oficial,sheet_name='Base Consolidada',dtype=D); N=base.copy()
    for c in D: N[c]=N[c].astype('string'); O[c]=O[c].astype('string').str.lstrip('0')
    for c in D: N[c]=N[c].str.lstrip('0')
    M=PPTO+MONTOS; filas=[]; det=[]
    niveles=[('Servicio × Tipo',['Servicio','Tipo']),('Programa × Subtítulo',['Servicio','Tipo','Programa','Subtítulo']),
             ('Programa × Asignación',['Servicio','Tipo','Programa','Subtítulo','Ítem','Asignación']),('Región',['Servicio','Tipo','Código'])]
    for nom,k in niveles:
        j=O.groupby(k,dropna=False)[M].sum().join(N.groupby(k,dropna=False)[M].sum(),rsuffix=' (agente)',how='outer').fillna(0)
        n=0
        for m in M:
            x=j[(j[m]-j[m+' (agente)']).abs()>1]; n+=len(x)
            for idx,r in x.iterrows():
                det.append({'Nivel':nom,'Clave':' | '.join(str(v) for v in (idx if isinstance(idx,tuple) else (idx,))),'Medida':m,'Oficial':r[m],'Agente':r[m+' (agente)'],'Diferencia':r[m+' (agente)']-r[m]})
        filas.append({'Nivel':nom,'Celdas comparadas':len(j)*len(M),'Con diferencia > 1 (miles de $)':n})
    GO=pd.read_excel(oficial,sheet_name='Base Glosas'); GM=['Ley de Presupuestos','Presupuesto Vigente']+MONTOS; k=['Servicio','Programa','Glosa']
    j=GO.groupby(k)[GM].sum().join(bg.groupby(k)[GM].sum(),rsuffix=' (agente)',how='outer').fillna(0); n=0
    for m in GM:
        x=j[(j[m]-j[m+' (agente)']).abs()>1]; n+=len(x)
        for idx,r in x.iterrows(): det.append({'Nivel':'Glosas','Clave':' | '.join(idx),'Medida':m,'Oficial':r[m],'Agente':r[m+' (agente)'],'Diferencia':r[m+' (agente)']-r[m]})
    filas.append({'Nivel':'Base Glosas: Servicio × Programa × Glosa','Celdas comparadas':len(j)*len(GM),'Con diferencia > 1 (miles de $)':n})
    return pd.DataFrame(filas), pd.DataFrame(det,columns=['Nivel','Clave','Medida','Oficial','Agente','Diferencia'])
