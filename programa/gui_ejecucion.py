"""Agente de Ejecución Presupuestaria – Power BI (ventana).
Se abre desde la ventana de bienvenida (inicio.py). Usa el mismo proceso que la versión de consola (proceso.py)."""
import os, sys, json, threading, queue, traceback, subprocess, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, font as tkfont

try:
    from proceso import *
    import actualizador as ACT
    import revision as RV
    ERROR_INICIO = None
except Exception as e:  # p.ej. faltan componentes
    ERROR_INICIO = e

# Paleta del framework gob.cl (Kit Digital, Gobierno de Chile)
AZUL = '#006fb3'; AZUL_OSC = '#0a132d'; AZUL_CLARO = '#a8b7c7'; ROJO = '#fe6565'; VERDE = '#2d717c'; AMARILLO = '#f2c728'
GRIS_FONDO = '#f8f9fa'; GRIS_BORDE = '#dee2e6'; GRIS_TEXTO = '#4a4a4a'; GRIS_SUAVE = '#6c757d'; BLANCO = '#ffffff'
RAIZ_AGENTE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# criticidad de cada cambio de versión: rojo (crítica: cambia o protege las cifras), amarillo (media: controles y
# procesos), verde (baja: ventana, documentos)
CRIT_COLOR = {'Crítica': '#b02a37', 'Media': '#b8860b', 'Baja': '#2e7d32'}
def separar_criticidad(texto):
    import re
    m = re.match(r'\[(Crítica|Media|Baja)\]\s*(.*)', str(texto), re.S)
    return (m.group(1), m.group(2)) if m else (None, str(texto))
RECURSOS = os.path.join(RAIZ_AGENTE, 'recursos')
HISTORIAL = os.path.join(RAIZ_AGENTE, 'Historial de versiones.pdf')


def abrir(path):
    try:
        if sys.platform.startswith('win'): os.startfile(path)
        elif sys.platform == 'darwin':
            r = subprocess.run(['/usr/bin/open', path], capture_output=True, text=True, timeout=20)
            if r.returncode != 0: raise OSError(r.stderr.strip() or f'open terminó con código {r.returncode}')
        else: subprocess.Popen(['xdg-open', path])
    except Exception as e:
        messagebox.showerror('No se pudo abrir', f'No se pudo abrir:\n{path}\n\n{e}')


def millones(m): return '$' + f'{m / 1000:,.0f}'.replace(',', '.')
def porc(x): return f'{x * 100:,.1f}'.replace('.', ',') + '%'


VIGILAR_MS = 3 * 60 * 1000   # cada cuánto se revisa si llegaron los informes del mes


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Agente de Ejecución Presupuestaria MINAGRI')
        self.geometry('1120x730'); self.minsize(1040, 690); self.configure(bg=GRIS_FONDO)
        self.cola = queue.Queue(); self.trabajando = False; self.resultado = None; self.estado = None; self.volver_al_inicio = False
        self._fuentes(); self._estilos(); self._icono()
        self._cabecera(); self._cuerpo(); self._pie()
        if ERROR_INICIO is not None:
            messagebox.showerror('Faltan componentes', 'El agente no pudo iniciar. Ejecute primero "1_INSTALAR.bat".\n\nDetalle: ' + str(ERROR_INICIO))
            self.destroy(); return
        self.after(100, self.cargar_config)
        self.after(150, self._procesar_cola)
        self.after(VIGILAR_MS, self.vigilar_informes)

    def report_callback_exception(self, exc, val, tb):
        """Un error en un botón o ventana no debe pasar en silencio (en Mac no hay consola): se muestra y se guarda en
        respaldos/errores.log para enviarlo."""
        texto = ''.join(traceback.format_exception(exc, val, tb))
        try:
            d = os.path.join(RAIZ_AGENTE, 'respaldos'); os.makedirs(d, exist_ok=True)
            with open(os.path.join(d, 'errores.log'), 'a', encoding='utf-8') as f:
                f.write(f'--- {datetime.datetime.now():%d-%m-%Y %H:%M:%S} · versión {VERSION}\n{texto}\n')
        except Exception:
            pass
        try:
            messagebox.showerror('Error en la ventana', f'{exc.__name__}: {val}\n\nEl detalle quedó en respaldos/errores.log del agente; '
                                 'envíelo a Manuel.')
        except Exception:
            pass

    # ---------------------------------------------------------------- apariencia
    def _fuentes(self):
        fam = set(tkfont.families())
        base = next((f for f in ('Roboto', 'Segoe UI', 'Arial', 'Helvetica') if f in fam), 'TkDefaultFont')
        tit = next((f for f in ('Roboto Slab', base) if f in fam), base)
        self.f = {'normal': (base, 10), 'chica': (base, 9), 'neg': (base, 10, 'bold'), 'paso': (tit, 13, 'bold'),
                  'titulo': (tit, 17, 'bold'), 'sub': (base, 10), 'boton': (base, 11, 'bold'), 'num': (base, 10)}

    def _estilos(self):
        s = ttk.Style(self)
        try: s.theme_use('clam')
        except tk.TclError: pass
        s.configure('.', font=self.f['normal'], background=BLANCO, foreground=GRIS_TEXTO)
        s.configure('Tarjeta.TFrame', background=BLANCO)
        s.configure('Fondo.TFrame', background=GRIS_FONDO)
        s.configure('TLabel', background=BLANCO, foreground=GRIS_TEXTO)
        s.configure('Fondo.TLabel', background=GRIS_FONDO)
        s.configure('Paso.TLabel', font=self.f['paso'], foreground=AZUL_OSC)
        s.configure('Suave.TLabel', foreground=GRIS_SUAVE, font=self.f['chica'])
        s.configure('Primario.TButton', font=self.f['boton'], foreground=BLANCO, background=AZUL, borderwidth=0, padding=(18, 10))
        s.map('Primario.TButton', background=[('disabled', AZUL_CLARO), ('active', '#005a92')])
        s.configure('Secundario.TButton', foreground=AZUL, background=BLANCO, bordercolor=AZUL, borderwidth=1, padding=(12, 6))
        s.map('Secundario.TButton', background=[('active', '#e6f1f8'), ('disabled', BLANCO)], foreground=[('disabled', AZUL_CLARO)])
        s.configure('Pie.TButton', foreground=AZUL, background=BLANCO, bordercolor=AZUL, borderwidth=1, padding=(8, 5))
        s.map('Pie.TButton', background=[('active', '#e6f1f8'), ('disabled', BLANCO)], foreground=[('disabled', AZUL_CLARO)])
        s.configure('TCombobox', padding=6); s.configure('TEntry', padding=6)
        s.configure('Treeview', rowheight=26, font=self.f['num'], background=BLANCO, fieldbackground=BLANCO, bordercolor=GRIS_BORDE)
        s.configure('Treeview.Heading', font=self.f['neg'], background=AZUL_OSC, foreground=BLANCO, relief='flat', padding=6)
        s.map('Treeview.Heading', background=[('active', AZUL_OSC)])
        s.configure('Horizontal.TProgressbar', troughcolor='#e9ecef', background=AZUL, bordercolor='#e9ecef', lightcolor=AZUL, darkcolor=AZUL, thickness=14)
        s.configure('TCheckbutton', background=BLANCO); s.map('TCheckbutton', background=[('disabled', BLANCO), ('active', BLANCO)])
        s.configure('TNotebook', background=BLANCO, borderwidth=0); s.configure('TNotebook.Tab', padding=(14, 6), font=self.f['neg'])
        s.map('TNotebook.Tab', background=[('selected', BLANCO), ('!selected', '#e9ecef')], foreground=[('selected', AZUL)])

    def _icono(self):
        for n in ('agente.png',):
            p = os.path.join(RECURSOS, n)
            if os.path.exists(p):
                try: self._ico = tk.PhotoImage(file=p); self.iconphoto(True, self._ico)
                except tk.TclError: pass
        ico = os.path.join(RECURSOS, 'agente.ico')
        if sys.platform.startswith('win') and os.path.exists(ico):
            try: self.iconbitmap(ico)
            except tk.TclError: pass

    def _cabecera(self):
        # franja institucional azul / rojo
        franja = tk.Frame(self, height=6, bg=AZUL); franja.pack(fill='x')
        tk.Frame(franja, height=6, width=260, bg=ROJO).pack(side='left')
        cab = tk.Frame(self, bg=BLANCO, padx=24, pady=10); cab.pack(fill='x')
        logo = os.path.join(RECURSOS, 'logo.png'); self._logo = None
        if os.path.exists(logo):
            try:
                img = tk.PhotoImage(file=logo)
                k = max(1, int(img.height() / 64 + 0.999)); self._logo = img.subsample(k, k)
                tk.Label(cab, image=self._logo, bg=BLANCO).pack(side='left', padx=(0, 18))
            except tk.TclError:
                self._logo = None
        if self._logo is None:
            bloque = tk.Frame(cab, bg=BLANCO); bloque.pack(side='left', padx=(0, 18))
            tk.Frame(bloque, width=6, height=52, bg=AZUL).pack(side='left'); tk.Frame(bloque, width=6, height=52, bg=ROJO).pack(side='left', padx=(0, 10))
            t = tk.Frame(bloque, bg=BLANCO); t.pack(side='left')
            tk.Label(t, text='Ministerio de', font=self.f['sub'], bg=BLANCO, fg=AZUL_OSC).pack(anchor='w')
            tk.Label(t, text='Agricultura', font=self.f['paso'], bg=BLANCO, fg=AZUL_OSC).pack(anchor='w')
            tk.Label(t, text='Gobierno de Chile', font=self.f['chica'], bg=BLANCO, fg=GRIS_SUAVE).pack(anchor='w')
        tk.Frame(cab, width=1, height=56, bg=GRIS_BORDE).pack(side='left', padx=(0, 18))
        txt = tk.Frame(cab, bg=BLANCO); txt.pack(side='left')
        tk.Label(txt, text='Agente de Ejecución Presupuestaria', font=self.f['titulo'], bg=BLANCO, fg=AZUL_OSC).pack(anchor='w')
        tk.Label(txt, text='UPROGE · Informe de Ejecución Mensual Power BI · Partida 13', font=self.f['sub'], bg=BLANCO, fg=GRIS_SUAVE).pack(anchor='w')
        tk.Frame(self, height=1, bg=GRIS_BORDE).pack(fill='x')

    def _tarjeta(self, padre, numero, titulo, **pack):
        f = tk.Frame(padre, bg=BLANCO, highlightbackground=GRIS_BORDE, highlightthickness=1, padx=16, pady=10)
        f.pack(fill='x', pady=(0, 10), **pack)
        h = tk.Frame(f, bg=BLANCO); h.pack(fill='x', pady=(0, 6))
        tk.Label(h, text=str(numero), font=self.f['neg'], bg=AZUL, fg=BLANCO, width=2, pady=2).pack(side='left', padx=(0, 10))
        ttk.Label(h, text=titulo, style='Paso.TLabel').pack(side='left')
        return f

    def _cuerpo(self):
        cuerpo = ttk.Frame(self, style='Fondo.TFrame', padding=(20, 12, 20, 0)); cuerpo.pack(fill='both', expand=True)
        izq = ttk.Frame(cuerpo, style='Fondo.TFrame'); izq.pack(side='left', fill='y')
        der = ttk.Frame(cuerpo, style='Fondo.TFrame'); der.pack(side='left', fill='both', expand=True, padx=(16, 0))
        izq.configure(width=440); izq.pack_propagate(False)

        # Paso 1: mes
        t1 = self._tarjeta(izq, 1, 'Elegir el mes')
        fila = tk.Frame(t1, bg=BLANCO); fila.pack(fill='x')
        self.var_mes = tk.StringVar()
        self.cb_mes = ttk.Combobox(fila, textvariable=self.var_mes, state='readonly', width=26, font=self.f['normal'])
        self.cb_mes.pack(side='left'); self.cb_mes.bind('<<ComboboxSelected>>', lambda e: self.revisar_mes())
        ttk.Button(fila, text='Actualizar', style='Secundario.TButton', command=self.cargar_meses).pack(side='left', padx=8)
        self.lbl_carpeta = ttk.Label(t1, text='', style='Suave.TLabel', wraplength=380, justify='left'); self.lbl_carpeta.pack(anchor='w', pady=(6, 2))
        self.grid_arch = tk.Frame(t1, bg=BLANCO); self.grid_arch.pack(fill='x')
        self.lbl_oficial = ttk.Label(t1, text='', style='Suave.TLabel', wraplength=380, justify='left'); self.lbl_oficial.pack(anchor='w', pady=(4, 0))

        # Paso 2: fecha de corte
        t2 = self._tarjeta(izq, 2, 'Fecha de corte de decretos')
        fila2 = tk.Frame(t2, bg=BLANCO); fila2.pack(fill='x')
        self.var_corte = tk.StringVar()
        self.ent_corte = ttk.Entry(fila2, textvariable=self.var_corte, width=14, font=self.f['normal']); self.ent_corte.pack(side='left')
        ttk.Label(fila2, text='(DD-MM-AAAA)', style='Suave.TLabel').pack(side='left', padx=6)
        self.btn_sugerir = ttk.Button(fila2, text='Sugerir fecha', style='Secundario.TButton', command=self.sugerir)
        self.btn_sugerir.pack(side='left', padx=(8, 0))
        self.lbl_corte = ttk.Label(t2, text='', style='Suave.TLabel', wraplength=380, justify='left'); self.lbl_corte.pack(anchor='w', pady=(6, 0))
        self.var_guardar = tk.BooleanVar(value=True)
        self.chk_guardar = ttk.Checkbutton(t2, text='Guardar esta fecha en el Catálogo Maestro', variable=self.var_guardar)
        self.chk_guardar.pack(anchor='w', pady=(6, 0))

        # Paso 3: generar
        t3 = self._tarjeta(izq, 3, 'Generar el consolidado')
        self.btn_generar = ttk.Button(t3, text='Generar consolidado', style='Primario.TButton', command=self.generar)
        self.btn_generar.pack(fill='x')
        self.prog = ttk.Progressbar(t3, mode='determinate', maximum=26, style='Horizontal.TProgressbar'); self.prog.pack(fill='x', pady=(10, 4))
        self.lbl_prog = ttk.Label(t3, text='Listo para comenzar.', style='Suave.TLabel', wraplength=380); self.lbl_prog.pack(anchor='w')

        # Resultados
        tr = tk.Frame(der, bg=BLANCO, highlightbackground=GRIS_BORDE, highlightthickness=1, padx=16, pady=10); tr.pack(fill='both', expand=True, pady=(0, 10))
        h = tk.Frame(tr, bg=BLANCO); h.pack(fill='x')
        ttk.Label(h, text='Resultado', style='Paso.TLabel').pack(side='left')
        self.lbl_res = ttk.Label(h, text='Aún no se ha generado el consolidado.', style='Suave.TLabel'); self.lbl_res.pack(side='left', padx=12)
        self.nb = ttk.Notebook(tr); self.nb.pack(fill='both', expand=True, pady=(8, 8))
        # pestaña ejecución
        p1 = ttk.Frame(self.nb, style='Tarjeta.TFrame', padding=(0, 8)); self.nb.add(p1, text='Ejecución')
        cols = ('serv', 'vig', 'dev', 'pct')
        self.tv = ttk.Treeview(p1, columns=cols, show='headings', height=8)
        for c, t, w, a in (('serv', 'Servicio', 140, 'w'), ('vig', 'Vigente ($ MM)', 130, 'e'), ('dev', 'Devengado ($ MM)', 140, 'e'), ('pct', '% avance', 90, 'e')):
            self.tv.heading(c, text=t, anchor=a); self.tv.column(c, width=w, minwidth=70, anchor=a, stretch=True)
        self.tv.tag_configure('total', font=self.f['neg'], background='#e6f1f8'); self.tv.tag_configure('par', background=GRIS_FONDO)
        self.tv.pack(fill='both', expand=True)
        # pestaña avisos
        p2 = ttk.Frame(self.nb, style='Tarjeta.TFrame', padding=(0, 8)); self.nb.add(p2, text='Avisos')
        self.txt_av = tk.Text(p2, wrap='word', font=self.f['normal'], relief='flat', bg=BLANCO, fg=GRIS_TEXTO, padx=6, pady=6, height=10)
        self.txt_av.pack(fill='both', expand=True)
        bav = tk.Frame(p2, bg=BLANCO); bav.pack(fill='x', pady=(6, 0))
        self.btn_cat = ttk.Button(bav, text='Agregar al catálogo…', style='Secundario.TButton', command=self.agregar_al_catalogo)
        self.btn_cat.pack(side='left'); self.btn_cat.state(['disabled'])
        self.lbl_cat = ttk.Label(bav, text='', style='Suave.TLabel'); self.lbl_cat.pack(side='left', padx=8)
        self.txt_av.tag_configure('rev', foreground='#b02a37', font=self.f['neg']); self.txt_av.tag_configure('info', foreground=GRIS_SUAVE)
        self.txt_av.tag_configure('ok', foreground=VERDE, font=self.f['neg'])
        # pestaña revisión
        pr = ttk.Frame(self.nb, style='Tarjeta.TFrame', padding=(0, 6)); self.nb.add(pr, text='Revisión')
        self.lbl_rev = ttk.Label(pr, text='Se completa al generar el consolidado.', style='Suave.TLabel', wraplength=560)
        self.lbl_rev.pack(anchor='w', pady=(0, 4))
        fr = tk.Frame(pr, bg=BLANCO); fr.pack(fill='both', expand=True)
        cols = ('niv', 'ctl', 'srv', 'det', 'est')
        self.tv_rev = ttk.Treeview(fr, columns=cols, show='headings', height=6, selectmode='extended')
        for c, t, w, st in (('niv', '', 24, False), ('ctl', 'Control', 150, False), ('srv', 'Servicio', 105, False),
                            ('det', 'Detalle', 190, True), ('est', 'Estado', 90, False)):
            self.tv_rev.heading(c, text=t, anchor='w'); self.tv_rev.column(c, width=w, minwidth=w if not st else 120, anchor='w', stretch=st)
        sb = ttk.Scrollbar(fr, orient='vertical', command=self.tv_rev.yview); self.tv_rev.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y'); self.tv_rev.pack(side='left', fill='both', expand=True)
        self.tv_rev.tag_configure('Rojo', foreground='#b02a37'); self.tv_rev.tag_configure('Amarillo', foreground='#8a6d00')
        self.tv_rev.tag_configure('ok', foreground=GRIS_SUAVE)
        self.tv_rev.bind('<Double-1>', lambda e: self.ver_alerta())
        br = tk.Frame(pr, bg=BLANCO); br.pack(fill='x', pady=(6, 0))
        self.btn_aceptar = ttk.Button(br, text='Aceptar…', style='Secundario.TButton', command=self.aceptar_alertas); self.btn_aceptar.pack(side='left')
        self.btn_quitar = ttk.Button(br, text='Quitar', width=7, style='Secundario.TButton', command=self.quitar_aceptacion); self.btn_quitar.pack(side='left', padx=6)
        self.btn_rev_xls = ttk.Button(br, text='Abrir revisión', style='Secundario.TButton', command=lambda: self.resultado and abrir(self.resultado['ruta_revision'])); self.btn_rev_xls.pack(side='left')
        def abrir_registro():
            import decretos as DC
            if os.path.exists(DC.REGISTRO): abrir(DC.REGISTRO)
            else: messagebox.showinfo('Registro de decretos', 'El registro se crea la primera vez que se genera un mes (al verificar las fechas en DIPRES).')
        ttk.Button(br, text='Decretos', width=9, style='Secundario.TButton', command=abrir_registro).pack(side='left', padx=6)
        self.btn_hoja = ttk.Button(br, text='Hoja de control', style='Secundario.TButton', command=self.exportar_hoja); self.btn_hoja.pack(side='right')
        for w in (self.btn_aceptar, self.btn_quitar, self.btn_rev_xls, self.btn_hoja): w.state(['disabled'])
        # pestaña validación del corte (semáforo de precisión + revisión del Power BI)
        pv = ttk.Frame(self.nb, style='Tarjeta.TFrame', padding=(0, 6)); self.nb.add(pv, text='Validación')
        self.lbl_vc = ttk.Label(pv, text='Se completa al generar el consolidado.', style='Suave.TLabel', wraplength=560); self.lbl_vc.pack(anchor='w', pady=(0, 4))
        fv = tk.Frame(pv, bg=BLANCO); fv.pack(fill='both', expand=True)
        self.tv_vc = ttk.Treeview(fv, columns=('est', 'ctl', 'res'), show='headings', height=6)
        for c, t, w, st in (('est', '', 24, False), ('ctl', 'Control', 210, False), ('res', 'Resultado', 300, True)):
            self.tv_vc.heading(c, text=t, anchor='w'); self.tv_vc.column(c, width=w, minwidth=w if not st else 160, anchor='w', stretch=st)
        sbv = ttk.Scrollbar(fv, orient='vertical', command=self.tv_vc.yview); self.tv_vc.configure(yscrollcommand=sbv.set)
        sbv.pack(side='right', fill='y'); self.tv_vc.pack(side='left', fill='both', expand=True)
        for k, c in (('Rojo', '#b02a37'), ('Amarillo', '#8a6d00'), ('Verde', '#2e7d32')): self.tv_vc.tag_configure(k, foreground=c)
        self.tv_vc.tag_configure('grupo', font=self.f['neg'])
        self.tv_vc.bind('<Double-1>', lambda e: self.ver_validacion())
        bv = tk.Frame(pv, bg=BLANCO); bv.pack(fill='x', pady=(6, 0))
        self.btn_val_xls = ttk.Button(bv, text='Abrir validación', style='Secundario.TButton', command=lambda: self.resultado and self.resultado.get('ruta_validacion') and abrir(self.resultado['ruta_validacion']))
        self.btn_val_xls.pack(side='left')
        self.btn_pbix = ttk.Button(bv, text='Revisar informe Power BI…', style='Secundario.TButton', command=self.revisar_pbix); self.btn_pbix.pack(side='left', padx=6)
        for w in (self.btn_val_xls, self.btn_pbix): w.state(['disabled'])
        # pestaña comparación con el oficial
        p3 = ttk.Frame(self.nb, style='Tarjeta.TFrame', padding=(0, 8)); self.nb.add(p3, text='Comparación')
        self.tv_val = ttk.Treeview(p3, columns=('niv', 'n', 'd'), show='headings', height=6)
        for c, t, w, a in (('niv', 'Nivel comparado', 300, 'w'), ('n', 'Cifras comparadas', 140, 'e'), ('d', 'Con diferencia', 120, 'e')):
            self.tv_val.heading(c, text=t, anchor=a); self.tv_val.column(c, width=w, anchor=a)
        self.tv_val.pack(fill='both', expand=True)
        self.lbl_val = ttk.Label(p3, text='Se muestra solo cuando ya existe el consolidado oficial del mes.', style='Suave.TLabel', wraplength=520)
        self.lbl_val.pack(anchor='w', pady=(6, 0))
        # pestaña registro
        p4 = ttk.Frame(self.nb, style='Tarjeta.TFrame', padding=(0, 8)); self.nb.add(p4, text='Registro')
        self.txt_log = tk.Text(p4, wrap='word', font=self.f['chica'], relief='flat', bg=GRIS_FONDO, fg=GRIS_TEXTO, padx=6, pady=6, height=10)
        self.txt_log.pack(fill='both', expand=True)
        # botones de resultado (dos filas, para que quepan con la ventana en su tamaño mínimo)
        b = tk.Frame(tr, bg=BLANCO); b.pack(fill='x')
        self.btn_abrir = ttk.Button(b, text='Abrir Excel', style='Primario.TButton', command=lambda: self.resultado and abrir(self.resultado['salida']))
        self.btn_abrir.pack(side='left')
        self.btn_ideas = ttk.Button(b, text='Ideas Fuerza', style='Secundario.TButton', command=self.ver_ideas); self.btn_ideas.pack(side='left', padx=8)
        self.btn_resumen = ttk.Button(b, text='Resumen ejecutivo', style='Secundario.TButton', command=self.resumen_ejecutivo); self.btn_resumen.pack(side='left', padx=(0, 8))
        self.btn_carpeta = ttk.Button(b, text='Carpeta del mes', style='Secundario.TButton', command=lambda: self.estado and abrir(os.path.join(self.estado['carpeta'], 'MINAGRI')))
        self.btn_carpeta.pack(side='left')
        b2 = tk.Frame(tr, bg=BLANCO); b2.pack(fill='x', pady=(8, 0))
        tk.Label(b2, text='Informes HTML:', font=self.f['chica'], bg=BLANCO, fg=GRIS_SUAVE).pack(side='left', padx=(0, 8))
        self.btn_dash = ttk.Button(b2, text='Generar Dashboard', style='Secundario.TButton', command=self.generar_dashboard)
        self.btn_dash.pack(side='left', padx=(0, 8))
        self.btn_mens = ttk.Button(b2, text='Reporte Mensualizado', style='Secundario.TButton', command=self.reporte_mensualizado)
        self.btn_mens.pack(side='left', padx=(0, 8))
        self.btn_publicar = ttk.Button(b2, text='Aprobar y publicar…', style='Secundario.TButton', command=self.aprobar_publicar)
        self.btn_publicar.pack(side='right'); self.btn_publicar.state(['disabled'])
        for w in (self.btn_abrir, self.btn_ideas, self.btn_resumen, self.btn_dash, self.btn_mens): w.state(['disabled'])

    def _pie(self):
        pie = tk.Frame(self, bg=GRIS_FONDO, padx=20, pady=8); pie.pack(fill='x', side='bottom')
        ttk.Button(pie, text='‹ Inicio', style='Pie.TButton', width=8, command=self.ir_al_inicio).pack(side='left', padx=(0, 8))
        ttk.Button(pie, text='Configurar carpetas', style='Pie.TButton', command=self.configurar).pack(side='left')
        ttk.Button(pie, text='Abrir Catálogo Maestro', style='Pie.TButton', command=lambda: abrir(self.cfg['catalogo'])).pack(side='left', padx=6)
        man = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'Manual de usuario.pdf')
        if os.path.exists(man): ttk.Button(pie, text='Manual de usuario', style='Pie.TButton', command=lambda: abrir(man)).pack(side='left')
        ttk.Button(pie, text='Actualizaciones', style='Pie.TButton', command=lambda: Actualizaciones(self)).pack(side='left', padx=6)
        ttk.Button(pie, text='Historial de versiones', style='Pie.TButton', command=lambda: HistorialVersiones(self)).pack(side='left')
        ttk.Button(pie, text='Panel de análisis', style='Pie.TButton', command=self.panel_analisis).pack(side='left', padx=6)
        tk.Label(pie, text=f'Versión {VERSION if ERROR_INICIO is None else "?"}' + (' beta' if ERROR_INICIO is None and ACT.es_beta(VERSION) else ''),
                 font=self.f['chica'], bg=GRIS_FONDO, fg=GRIS_SUAVE).pack(side='right')

    # ---------------------------------------------------------------- configuración y mes
    def cargar_config(self):
        try:
            self.cfg = leer_config()
        except ErrorAgente as e:
            messagebox.showerror('Configuración', str(e)); return
        prob = revisar_rutas(self.cfg)
        if prob:
            messagebox.showwarning('Configurar carpetas', 'Antes de comenzar, revise las carpetas:\n\n• ' + '\n• '.join(prob))
            self.configurar(); return
        self.cargar_meses()
        if not getattr(self, '_act_revisada', False):
            self._act_revisada = True; self.after(1200, self.revisar_actualizacion)

    def vigilar_informes(self):
        """Con la ventana abierta, avisa cuando el mes más reciente sin consolidado del agente ya tiene sus 12 informes."""
        try:
            if not self.trabajando and getattr(self, 'cfg', None):
                meses = meses_disponibles(self.cfg['carpeta_informes'])
                avisados = getattr(self, '_avisados', None)
                if avisados is None:      # primera vuelta: se recuerda lo que ya estaba completo, sin avisar
                    self._avisados = {d for a, m, d in meses if informes_completos(os.path.join(self.cfg['carpeta_informes'], d))}
                else:
                    for a, m, d in meses:
                        c = os.path.join(self.cfg['carpeta_informes'], d)
                        if d in avisados or os.path.exists(salida_del_mes(c, a, m)) or not informes_completos(c): continue
                        avisados.add(d)
                        if [x for x in meses if x[2] == d] and messagebox.askyesno('Informes completos',
                                f'Ya están los 12 informes de {MESES[m-1]} {a}.\n\n¿Seleccionar el mes para generar el consolidado?'):
                            self.cargar_meses()
                            i = next((k for k, x in enumerate(self.meses) if x[2] == d), None)
                            if i is not None: self.cb_mes.current(i); self.revisar_mes()
                        break
        except Exception:
            pass
        self.after(VIGILAR_MS, self.vigilar_informes)

    def ir_al_inicio(self):
        if self.trabajando:
            messagebox.showinfo('Espere', 'Hay un consolidado en proceso. Vuelva al inicio cuando termine.'); return
        self.volver_al_inicio = True; self.destroy()

    def destroy(self):
        # cancela las tareas programadas (revisión de la cola, avisos) para que no queden pendientes al volver al inicio
        try:
            for i in self.tk.splitlist(self.tk.call('after', 'info')): self.after_cancel(i)
        except tk.TclError:
            pass
        super().destroy()

    # ---------------------------------------------------------------- actualizaciones
    def revisar_actualizacion(self):
        """Al abrir: si hay una versión nueva (carpeta compartida o GitHub), ofrece instalarla una vez por sesión."""
        carpeta = self.cfg.get('carpeta_actualizaciones', '')
        repo = ACT.repositorio_github(self.cfg.get('repositorio_github', ''))
        if (not carpeta and not repo) or ACT.REVISADO[0]: return   # la bienvenida ya revisó en esta sesión
        def buscar():
            info = None
            try: info = ACT.buscar_en_carpeta(carpeta, VERSION) if carpeta else None
            except Exception: info = None
            if info is None and repo:
                try: info = ACT.buscar_en_github(repo, VERSION)
                except Exception: info = None      # sin conexión o GitHub bloqueado: no se molesta al usuario
            if info and info['version'] not in ACT.YA_OFRECIDAS:
                ACT.YA_OFRECIDAS.add(info['version'])
                hallada[0] = info
            hallada[1] = True
        hallada = [None, False]
        def esperar(n=0):                 # la ventana solo se toca desde el hilo principal (macOS)
            if hallada[1]:
                if hallada[0]: self.ofrecer_actualizacion(hallada[0])
            elif n < 120:
                self.after(500, lambda: esperar(n + 1))
        threading.Thread(target=buscar, daemon=True).start()
        esperar()

    def ofrecer_actualizacion(self, info):
        if self.trabajando: return
        cambios = '\n'.join('• ' + c for c in info.get('cambios', [])[:8])
        if messagebox.askyesno('Hay una versión nueva del agente',
                               f'Está disponible la versión {info["version"]} ({info.get("fecha", "")}); usted tiene la {VERSION}.\n\n'
                               f'{cambios}\n\n¿Instalarla ahora? Su configuración y el Catálogo Maestro se conservan, '
                               'y la versión actual queda respaldada.'):
            self.instalar_actualizacion(info)

    def instalar_actualizacion(self, info, ventana=None):
        """info: dict de inspeccionar()/buscar_en_github(), o la ruta de un .zip."""
        if self.trabajando:
            messagebox.showinfo('Espere', 'Hay un consolidado en proceso. Instale la actualización cuando termine.', parent=ventana or self); return
        try:
            self.configure(cursor='watch'); self.update()
            if isinstance(info, str): ruta = info
            elif info.get('url'): ruta = ACT.descargar(info)
            else: ruta = info['archivo']
            gh = not isinstance(info, str) and bool(info.get('url'))
            r = ACT.instalar(ruta, log=lambda t: None, origen='github' if gh else 'archivo',
                             confirmar_sin_firma=lambda d: messagebox.askyesno('Sin firma digital', 'Esta actualización no trae firma digital, por lo que el agente no puede comprobar que venga de UPROGE.\n\nInstálela solo si el .zip se lo entregó UPROGE. ¿Instalar igual?', icon='warning', parent=ventana or self))
        except Exception as e:
            self.configure(cursor='')
            messagebox.showerror('No se pudo actualizar', f'{e}\n\nNo se modificó config.ini ni el Catálogo Maestro.', parent=ventana or self); return
        self.configure(cursor='')
        messagebox.showinfo('Actualización instalada', f'Versión {r["version"]} instalada (antes: {r["anterior"]}).\n\n'
                            + '\n'.join('• ' + h for h in r['hechos']) + '\n\nLa ventana se cerrará y se abrirá de nuevo con la versión nueva.', parent=ventana or self)
        self.reiniciar()

    def reiniciar(self):
        ejecutable = sys.executable
        if sys.platform.startswith('win') and ejecutable.lower().endswith('python.exe'):
            w = ejecutable[:-10] + 'pythonw.exe'
            if os.path.exists(w): ejecutable = w
        try:
            subprocess.Popen([ejecutable, os.path.join(RAIZ_AGENTE, 'programa', 'ventana.pyw')], cwd=os.path.join(RAIZ_AGENTE, 'programa'))
        except Exception as e:
            messagebox.showwarning('Abra el agente de nuevo', f'No se pudo reabrir automáticamente ({e}). Ábralo desde el acceso directo.')
        self.destroy()

    def cargar_meses(self):
        self.meses = meses_disponibles(self.cfg['carpeta_informes'])
        self.cb_mes['values'] = [f'{MESES[m-1]} {a}' for a, m, _ in self.meses]
        if self.meses:
            # por defecto: el mes más reciente que ya tiene los 12 informes
            self.cb_mes.current(0)
            for i, (a, m, d) in enumerate(self.meses):
                try:
                    ubicar_archivos(os.path.join(self.cfg['carpeta_informes'], d)); self.cb_mes.current(i); break
                except Exception: continue
            self.revisar_mes()
        else:
            self.lbl_carpeta.configure(text='No hay carpetas de meses ("MM Mes AAAA") en la carpeta de informes configurada.')

    def mes_sel(self):
        i = self.cb_mes.current()
        return (self.meses[i][0], self.meses[i][1]) if i >= 0 else (None, None)

    def revisar_mes(self):
        anio, m = self.mes_sel()
        if not anio: return
        for w in self.grid_arch.winfo_children(): w.destroy()
        try:
            self.estado = estado_del_mes(self.cfg, anio, m)
        except Exception as e:
            self.lbl_carpeta.configure(text=str(e)); self.estado = None; return
        est = self.estado
        self.lbl_carpeta.configure(text='Carpeta: …\\' + os.path.basename(est['carpeta']))
        from motor import SERV
        tk.Label(self.grid_arch, text='Informes del mes (Gastos e Ingresos):', bg=BLANCO, fg=GRIS_SUAVE, font=self.f['chica']).grid(row=0, column=0, columnspan=3, sticky='w')
        for i, s in enumerate(SERV):
            falta = [t for t in ('Gastos', 'Ingresos') if (s, t) not in est['archivos']]
            c = tk.Frame(self.grid_arch, bg=BLANCO); c.grid(row=1 + i // 3, column=i % 3, sticky='w', padx=(0, 14), pady=1)
            tk.Label(c, text='✖' if falta else '✔', fg='#b02a37' if falta else VERDE, bg=BLANCO, font=self.f['neg']).pack(side='left')
            tk.Label(c, text=' ' + s.split()[1] + (f' (falta {" e ".join(f.lower() for f in falta)})' if falta else ''),
                     fg='#b02a37' if falta else GRIS_TEXTO, bg=BLANCO, font=self.f['chica']).pack(side='left')
        # revisión previa de los insumos (errores impiden generar)
        prev = est.get('previa', []); errores = [t for n, t in prev if n == 'Error']; avisos = [t for n, t in prev if n != 'Error']
        fila = len(SERV) // 3 + 2
        for t in errores[:3]:
            tk.Label(self.grid_arch, text='✖ ' + t, fg='#b02a37', bg=BLANCO, font=self.f['chica'], wraplength=390, justify='left'
                     ).grid(row=fila, column=0, columnspan=3, sticky='w', pady=(3, 0)); fila += 1
        resto = errores[3:] + avisos
        if resto:
            l = tk.Label(self.grid_arch, text=('⚠ ' if avisos else '✖ ') + f'{len(resto)} observación(es) de la revisión previa · ver detalle', fg='#8a6d00',
                         bg=BLANCO, font=self.f['chica'] + ('underline',), cursor='hand2')
            l.grid(row=fila, column=0, columnspan=3, sticky='w', pady=(3, 0)); fila += 1
            l.bind('<Button-1>', lambda e, r=resto: messagebox.showinfo('Revisión previa de los insumos', '• ' + '\n\n• '.join(r)))
        if not prev:
            tk.Label(self.grid_arch, text='✔ Revisión previa sin observaciones', fg=VERDE, bg=BLANCO, font=self.f['chica']).grid(row=fila, column=0, columnspan=3, sticky='w', pady=(3, 0))
        if est['problemas_archivos'] or errores:
            self.lbl_oficial.configure(text='Corrija lo marcado con ✖ y presione "Actualizar".')
        else:
            self.lbl_oficial.configure(text=('Se comparará con el consolidado oficial del mes.' if est['oficial'] else 'Sin consolidado oficial del mes: no habrá comparación.')
                                       + (' Ya existe un archivo AGENTE; se reemplazará.' if os.path.exists(est['salida']) else ''))
        # fecha de corte
        if est['corte_catalogo'] is not None:
            self.var_corte.set(est['corte_catalogo'].strftime('%d-%m-%Y'))
            self.lbl_corte.configure(text='Fecha registrada en el Catálogo Maestro.'); self.var_guardar.set(False); self.chk_guardar.state(['disabled'])
        else:
            self.var_corte.set(''); self.corte_sugerida = None; self.chk_guardar.state(['!disabled']); self.var_guardar.set(True)
            self.lbl_corte.configure(text='El catálogo no tiene fecha para este mes. ' + ('Use "Sugerir fecha" o escríbala.' if est['oficial'] else 'Escríbala (fecha hasta la cual se consideran los decretos).'))
        self.btn_sugerir.state(['!disabled'] if est['oficial'] else ['disabled'])
        self.btn_generar.state(['disabled'] if (est['problemas_archivos'] or errores) else ['!disabled'])
        self.btn_publicar.state(['!disabled'] if os.path.exists(est['salida']) else ['disabled'])

    def sugerir(self):
        if not self.estado or not self.estado['oficial']: return
        self.lbl_corte.configure(text='Calculando la fecha que reproduce el vigente del consolidado oficial…'); self.update_idletasks()
        try:
            d, h, e = sugerir_corte_mes(self.cfg, self.estado['oficial'])
        except Exception as ex:
            self.lbl_corte.configure(text=f'No se pudo calcular: {ex}'); return
        self.var_corte.set(d.strftime('%d-%m-%Y')); self.corte_sugerida = d
        self.lbl_corte.configure(text=(f'Sugerida: reproduce el vigente oficial con cualquier fecha entre {d:%d-%m-%Y} y {h:%d-%m-%Y}.' if h is not None else f'Sugerida: {d:%d-%m-%Y}.')
                                 + ('' if e == 0 else f' Ojo: {e} servicio(s) no cuadran.'))

    def configurar(self):
        Configuracion(self)

    # ---------------------------------------------------------------- generar
    def generar(self):
        if self.trabajando or not self.estado: return
        anio, m = self.mes_sel()
        try:
            corte = leer_fecha(self.var_corte.get())
        except ErrorAgente as e:
            messagebox.showwarning('Fecha de corte', str(e) if self.var_corte.get().strip() else 'Falta la fecha de corte de decretos (paso 2).'); return
        if self.estado['corte_catalogo'] is not None and corte == self.estado['corte_catalogo']: origen = 'hoja "Corte decretos" del catálogo'
        elif getattr(self, 'corte_sugerida', None) is not None and corte == self.corte_sugerida: origen = 'sugerida a partir del consolidado oficial'
        else: origen = 'ingresada por el usuario'
        try:
            import bitacora as BT
            hc = BT.ultima_hoja_de_control(self.cfg['carpeta_informes'], anio, m)
        except Exception:
            hc = None
        if hc and not messagebox.askyesno('Mes con hoja de control', f'{MESES[m-1]} {anio} ya tiene hoja de control exportada el {hc[0]} por {hc[1]}.\n\n'
                                          'Si vuelve a generar, las cifras pueden cambiar respecto de lo revisado. ¿Generar igual?'):
            return
        if os.path.exists(self.estado['salida']) and not messagebox.askyesno('Reemplazar', f'Ya existe:\n{os.path.basename(self.estado["salida"])}\n\n¿Desea reemplazarlo?'):
            return
        self.trabajando = True; self.resultado = None
        for w in (self.btn_generar, self.btn_abrir, self.btn_ideas, self.btn_resumen, self.cb_mes, self.btn_sugerir): w.state(['disabled'])
        self.prog['value'] = 0; self.txt_log.delete('1.0', 'end'); self.lbl_res.configure(text='Procesando…')
        guardar = self.var_guardar.get() and self.estado['corte_catalogo'] is None
        threading.Thread(target=self._trabajo, args=(anio, m, corte, origen, guardar), daemon=True).start()

    def _trabajo(self, anio, m, corte, origen, guardar):
        try:
            r = procesar(self.cfg, anio, m, corte, origen, self.estado['carpeta'],
                         log=lambda t: self.cola.put(('log', t)), paso=lambda n, tot, t: self.cola.put(('paso', n, t)))
            if guardar:
                guardar_corte_en_catalogo(self.cfg['catalogo'], anio, m, corte, f'Guardada desde la ventana el {datetime.date.today():%d-%m-%Y} ({origen}).')
                r['avisos'].append(('Info', f'Fecha de corte {corte:%d-%m-%Y} guardada en el Catálogo Maestro para {MESES[m-1]} {anio}.'))
            self.cola.put(('fin', r))
        except ErrorAgente as e:
            self.cola.put(('error', str(e), None))
        except Exception as e:
            self.cola.put(('error', f'Error inesperado: {e}', traceback.format_exc()))

    def _procesar_cola(self):
        try:
            while True:
                msg = self.cola.get_nowait()
                if msg[0] == 'log':
                    self.txt_log.insert('end', f'{datetime.datetime.now():%H:%M:%S}  {msg[1]}\n'); self.txt_log.see('end')
                elif msg[0] == 'paso':
                    self.prog['value'] = msg[1]; self.lbl_prog.configure(text=msg[2])
                elif msg[0] == 'fin':
                    self._mostrar(msg[1])
                elif msg[0] == 'error':
                    self._fin_trabajo(); self.lbl_prog.configure(text='No se pudo generar. Revise el mensaje.'); self.lbl_res.configure(text='Error')
                    if msg[2]: self.txt_log.insert('end', msg[2])
                    messagebox.showerror('No se pudo generar', msg[1] + ('\n\nEl detalle técnico quedó en la pestaña "Registro"; envíelo a Manuel.' if msg[2] else ''))
        except queue.Empty:
            pass
        self.after(120, self._procesar_cola)

    def _fin_trabajo(self):
        self.trabajando = False
        for w in (self.btn_generar, self.cb_mes): w.state(['!disabled'])
        if self.estado and self.estado['oficial']: self.btn_sugerir.state(['!disabled'])

    def _mostrar(self, r):
        self._fin_trabajo(); self.resultado = r
        for w in (self.btn_abrir, self.btn_ideas, self.btn_resumen, self.btn_dash, self.btn_mens, self.btn_publicar): w.state(['!disabled'])
        self.prog['value'] = self.prog['maximum']
        self.lbl_prog.configure(text='Listo. Archivo: ' + os.path.basename(r['salida']))
        self.lbl_res.configure(text=f'{r["nombre"]} · corte de decretos {r["corte"]:%d-%m-%Y}')
        # tabla
        self.tv.delete(*self.tv.get_children())
        g = r['ctrl'][r['ctrl'].Tipo == 'Gastos']
        for i, (_, x) in enumerate(g.iterrows()):
            self.tv.insert('', 'end', values=(x.Servicio, millones(x['Presupuesto Vigente']), millones(x['Monto Devengo']), porc(x['Monto Devengo'] / x['Presupuesto Vigente'])), tags=('par',) if i % 2 else ())
        v, d = g['Presupuesto Vigente'].sum(), g['Monto Devengo'].sum()
        self.tv.insert('', 'end', values=('TOTAL MINAGRI', millones(v), millones(d), porc(d / v)), tags=('total',))
        # avisos
        self.txt_av.delete('1.0', 'end')
        rev = [t for k, t in r['avisos'] if k == 'Revisar']; inf = [t for k, t in r['avisos'] if k != 'Revisar']
        if not rev: self.txt_av.insert('end', '✔ No hay avisos para revisar.\n\n', 'ok')
        for t in rev: self.txt_av.insert('end', '⚠ ' + t + '\n\n', 'rev')
        for t in inf: self.txt_av.insert('end', 'ℹ ' + t + '\n\n', 'info')
        self.nb.tab(1, text=f'Avisos ({len(rev)})' if rev else 'Avisos')
        pend = r.get('pendientes') or []
        self.btn_cat.state(['!disabled'] if pend else ['disabled'])
        self.lbl_cat.configure(text=f'{len(pend)} elemento(s) sin registrar en el Catálogo Maestro' if pend else 'El Catálogo Maestro tiene todo lo que trae el mes')
        # validación
        self.tv_val.delete(*self.tv_val.get_children())
        if r['validacion'] is not None:
            for _, x in r['validacion'].iterrows():
                self.tv_val.insert('', 'end', values=(x['Nivel'], f'{int(x["Celdas comparadas"]):,}'.replace(',', '.'), int(x['Con diferencia > 1 (miles de $)'])))
            self.lbl_val.configure(text='Diferencias mayores a 1 (miles de pesos). El detalle de cada una está en la hoja "Validación" del Excel generado.')
        else:
            self.lbl_val.configure(text='No hay consolidado oficial del mes, por lo que no se hizo la comparación.')
        self.mostrar_revision()
        self.mostrar_validacion()
        rs = RV.resumen(r['revision']) if r.get('revision') is not None else None
        self.nb.select(2 if rs and rs['rojos_pendientes'] else (1 if rev else 0))
        messagebox.showinfo('Listo', f'Se generó el consolidado de {r["nombre"]}.\n\n{os.path.basename(r["salida"])}\n\n'
                            + (f'Hay {len(rev)} aviso(s) para revisar.' if rev else 'No hay avisos para revisar.')
                            + (f'\n\nRevisión: {rs["rojos_pendientes"]} rojo(s) sin explicar y {rs["amarillos_pendientes"]} amarillo(s).' if rs else ''))

    # ---------------------------------------------------------------- panel de análisis (base histórica)
    def panel_analisis(self):
        import historia as HI
        try:
            n = len(HI.meses())
            if messagebox.askyesno('Panel de análisis', f'La base histórica de este equipo tiene {n} mes(es).\n\n'
                                   '¿Actualizarla ahora desde los consolidados de las carpetas de los meses (año actual y anterior)? '
                                   'Toma unos segundos por mes.\n\n(Con "No" se abre el panel con lo que ya tiene.)', parent=self):
                self.configure(cursor='watch'); self.update()
                HI.reconstruir([self.cfg.get('carpeta_informes', ''), self.cfg.get('carpeta_informes_anio_anterior', '')], log=lambda t: None)
            destino = os.path.join(self.cfg.get('carpeta_informes') or os.path.join(RAIZ_AGENTE, 'respaldos'), 'Panel de análisis Agentes UPROGE.html')
            HI.panel(destino, VERSION)
        except Exception as e:
            self.configure(cursor=''); messagebox.showerror('Panel de análisis', f'No se pudo preparar el panel ({type(e).__name__}: {e})', parent=self); return
        self.configure(cursor=''); abrir(destino)

    # ---------------------------------------------------------------- publicación con doble control
    def aprobar_publicar(self):
        """Publica el consolidado del agente como consolidado oficial del mes (el que lee Power BI), con doble control."""
        import seguridad as SG
        est = self.estado; r = self.resultado
        if not est: return
        anio, m = self.mes_sel()
        rev = r['revision'] if r and r.get('anio') == anio and r.get('mes') == m else RV._leer(RV.ruta_revision(est['carpeta'], anio, m))
        e = SG.estado_publicacion(self.cfg, anio, m, est['carpeta'], est['salida'], rev)
        if not e['puede']:
            messagebox.showwarning('Aún no se puede publicar', '\n\n'.join('• ' + x for x in e['motivos']), parent=self); return
        g = e['generacion'] or {}
        destino = SG.ruta_oficial(est['carpeta'], anio, m)
        if not messagebox.askyesno('Aprobar y publicar', f'Se publicará el consolidado de {MESES[m - 1]} {anio} como consolidado oficial:\n\n{os.path.basename(destino)}\n\n'
                                   f'Generado por {g.get("Usuario", "?")} el {g.get("Fecha", "?")} (versión {g.get("Versión del agente", "?")}).\n'
                                   f'Aprueba: {SG.usuario()}.\n\nEs el archivo que lee el informe Power BI. La aprobación queda en la bitácora. ¿Publicar?', parent=self): return
        try:
            self.configure(cursor='watch'); self.update()
            ruta = SG.publicar(self.cfg, anio, m, est['carpeta'], est['salida'], rev,
                               confirmar_reemplazo=lambda: messagebox.askyesno('Ya publicado', 'El agente ya había publicado este mes. ¿Reemplazarlo por la versión actual?', parent=self))
        except Exception as ex:
            self.configure(cursor=''); messagebox.showwarning('No se publicó', str(ex), parent=self); return
        self.configure(cursor='')
        messagebox.showinfo('Publicado', f'Consolidado oficial publicado:\n\n{ruta}\n\nAbra el informe Power BI del mes y presione Actualizar.', parent=self)

    # ---------------------------------------------------------------- dashboard HTML
    def generar_dashboard(self):
        """Dashboard HTML autocontenido del mes, con las mismas páginas del informe Power BI. Toma el consolidado tal
        como está guardado (incluidas las Ideas Fuerza que se hayan editado en el Excel)."""
        r = self.resultado
        if not r or not os.path.exists(r['salida']):
            messagebox.showinfo('Generar Dashboard', 'Primero genere el consolidado del mes.', parent=self); return
        from proceso import _abierto_en_excel
        if _abierto_en_excel(r['salida']) and not messagebox.askyesno('Generar Dashboard', 'El consolidado está abierto en Excel. El dashboard usa la última versión guardada '
                                                                      '(guarde primero si editó las Ideas Fuerza). ¿Continuar?', parent=self): return
        try:
            self.configure(cursor='watch'); self.update()
            import dashboard as DB
            ruta = DB.generar(r['salida'], r['anio'], r['mes'], r['carpeta'], VERSION)
        except Exception as e:
            self.configure(cursor=''); messagebox.showerror('Generar Dashboard', f'No se pudo generar el dashboard ({type(e).__name__}: {e})', parent=self); return
        self.configure(cursor='')
        try:
            import bitacora as BT; BT.registrar_accion(self.cfg['carpeta_informes'], r['anio'], r['mes'], 'Dashboard HTML generado', os.path.basename(ruta))
        except Exception: pass
        if messagebox.askyesno('Dashboard listo', f'Se generó el dashboard del mes:\n\n{os.path.basename(ruta)}\n\nEstá en la carpeta MINAGRI del mes. Es un solo archivo: '
                               'se abre con doble clic en el navegador y se puede enviar por correo; funciona sin internet.\n\n¿Abrirlo ahora?', parent=self):
            abrir(ruta)

    # ---------------------------------------------------------------- reporte mensualizado HTML
    def log(self, texto):
        self.txt_log.insert('end', f'{datetime.datetime.now():%H:%M:%S}  {texto}\n'); self.txt_log.see('end')

    def reporte_mensualizado(self):
        """Ejecución mes a mes del año (Ingresos y Gastos), con los consolidados de cada mes ya guardados en las carpetas."""
        r = self.resultado
        if not r or not os.path.exists(r['salida']):
            messagebox.showinfo('Reporte Mensualizado', 'Primero genere el consolidado del mes.', parent=self); return
        from proceso import _abierto_en_excel
        if _abierto_en_excel(r['salida']) and not messagebox.askyesno('Reporte Mensualizado', 'El consolidado está abierto en Excel. El reporte usa la última versión guardada. ¿Continuar?', parent=self): return
        import mensualizado as MN
        cortes = MN.consolidados(r['carpeta'], r['anio'], r['mes'], actual=r['salida'])
        faltan = [MESES[m - 1] for m, ruta, f in cortes if f == 'falta']
        if faltan and not messagebox.askyesno('Reporte Mensualizado', 'No se encontró el consolidado de: ' + ', '.join(faltan) + '.\n\nEsos meses (y el devengo mensual del mes '
                                              'siguiente) quedarán sin información en el reporte. ¿Generarlo igual?', parent=self): return
        self.btn_mens.state(['disabled']); self.configure(cursor='watch')
        self.log(f'Reporte Mensualizado {MESES[r["mes"] - 1]} {r["anio"]}: leyendo los consolidados de cada mes (la primera vez puede tardar un par de minutos)…')
        res, cola = {}, queue.Queue()   # la ventana solo se toca desde el hilo principal
        def trabajo():
            try: res['ruta'], res['cortes'] = MN.generar(r['salida'], r['anio'], r['mes'], r['carpeta'], VERSION, log=cola.put)
            except Exception as e: res['error'] = e
            res['fin'] = True
        def revisar():
            while not cola.empty(): self.log(cola.get())
            if res.get('fin'): fin()
            else: self.after(250, revisar)
        def fin():
            self.btn_mens.state(['!disabled']); self.configure(cursor='')
            if 'error' in res:
                e = res['error']; self.log(f'  No se pudo generar: {e}')
                messagebox.showerror('Reporte Mensualizado', f'No se pudo generar el reporte ({type(e).__name__}: {e})', parent=self); return
            ruta = res['ruta']; self.log(f'  Listo: {os.path.basename(ruta)}')
            try:
                import bitacora as BT; BT.registrar_accion(self.cfg['carpeta_informes'], r['anio'], r['mes'], 'Reporte Mensualizado HTML generado', os.path.basename(ruta))
            except Exception: pass
            if messagebox.askyesno('Reporte listo', f'Se generó el reporte mensualizado:\n\n{os.path.basename(ruta)}\n\nEstá en la carpeta MINAGRI del mes. Es un solo archivo: '
                                   'se abre con doble clic en el navegador, se puede enviar por correo y funciona sin internet.\n\n¿Abrirlo ahora?', parent=self):
                abrir(ruta)
        threading.Thread(target=trabajo, daemon=True).start(); self.after(250, revisar)

    # ---------------------------------------------------------------- validación del corte
    def mostrar_validacion(self):
        r = self.resultado or {}; h = r.get('validacion_corte'); sem = r.get('semaforo')
        self.tv_vc.delete(*self.tv_vc.get_children()); self._vc = {}
        if h is None or sem is None:
            self.lbl_vc.configure(text='No se pudo completar la validación del corte (ver Avisos).'); self.nb.tab(3, text='Validación')
            return
        ico = {'Rojo': '●', 'Amarillo': '●', 'Verde': '✔'}
        for _, g in sem.iterrows():
            gi = self.tv_vc.insert('', 'end', values=(ico[g.Estado], g.Control, g.Resultado), tags=(g.Estado, 'grupo'), open=g.Estado != 'Verde')
            for i, x in h[h.Grupo == g.Control].iterrows():
                if x.Criticidad == 'Verde' and g.Estado != 'Verde': continue
                iid = self.tv_vc.insert(gi, 'end', values=(ico[x.Criticidad], '   ' + x.Control, x.Detalle), tags=(x.Criticidad,)); self._vc[iid] = x.to_dict()
        pb = r.get('pbix') or []
        if pb:
            peor = 'Rojo' if any(x['Criticidad'] == 'Rojo' for x in pb) else ('Amarillo' if any(x['Criticidad'] == 'Amarillo' and 'manual' not in x['Elemento'] for x in pb) else 'Verde')
            gi = self.tv_vc.insert('', 'end', values=(ico[peor], 'Informe Power BI (.pbix)', f'{sum(x["Criticidad"] == "Rojo" for x in pb)} rojo(s), {sum(x["Criticidad"] == "Amarillo" and "manual" not in x["Elemento"] for x in pb)} amarillo(s) y un recordatorio'), tags=(peor, 'grupo'), open=True)
            for x in pb:
                iid = self.tv_vc.insert(gi, 'end', values=(ico[x['Criticidad']], '   ' + x['Elemento'], x['Detalle']), tags=(x['Criticidad'],))
                self._vc[iid] = {'Grupo': 'Informe Power BI', 'Control': x['Elemento'], 'Criticidad': x['Criticidad'], 'Servicio': '', 'Detalle': x['Detalle'], 'Qué revisar': ''}
        nr = int((h.Criticidad == 'Rojo').sum()) + sum(x['Criticidad'] == 'Rojo' for x in pb); na = int((h.Criticidad == 'Amarillo').sum()) + sum(x['Criticidad'] == 'Amarillo' and 'manual' not in x['Elemento'] for x in pb)
        self.lbl_vc.configure(text=('Todo en verde: el consolidado cuadra con los insumos y respeta la estructura del informe Power BI.' if not nr and not na else
                                    f'{nr} rojo(s) y {na} amarillo(s). Los rojos del consolidado pasan a la pestaña Revisión y deben corregirse o explicarse; los del informe Power BI se corrigen en Power BI Desktop.')
                              + (' Revise también el informe Power BI.' if not pb else '') + '  Doble clic para ver el detalle.')
        self.nb.tab(3, text=f'Validación ({nr} rojos)' if nr else 'Validación ✔')
        for w in (self.btn_val_xls, self.btn_pbix): w.state(['!disabled'])

    def ver_validacion(self):
        sel = self.tv_vc.selection()
        if not sel or sel[0] not in self._vc: return
        x = self._vc[sel[0]]
        messagebox.showinfo('Validación del corte', f'{x["Criticidad"]} · {x["Grupo"]}\n{x["Control"]}\n\n{x["Detalle"]}' + (f'\n\nQué revisar: {x["Qué revisar"]}' if x.get('Qué revisar') else ''), parent=self)

    def revisar_pbix(self):
        r = self.resultado
        if not r: return
        ini = r['carpeta']
        ruta = filedialog.askopenfilename(parent=self, title='Elegir el informe Power BI del mes', initialdir=ini, filetypes=[('Power BI', '*.pbix')])
        if not ruta: return
        try:
            import proceso as PRC
            bg = pd.read_excel(r['salida'], sheet_name='Base Glosas')
            r['pbix'] = PRC.revisar_pbix_del_mes(r['carpeta'], r['anio'], r['mes'], r['base'], bg, [ruta])
        except Exception as e:
            messagebox.showerror('Informe Power BI', f'No se pudo revisar el archivo ({type(e).__name__}: {e})', parent=self); return
        self.mostrar_validacion(); self.nb.select(3)

    # ---------------------------------------------------------------- revisión
    def mostrar_revision(self):
        r = self.resultado; d = r.get('revision') if r else None
        self.tv_rev.delete(*self.tv_rev.get_children())
        if d is None:
            self.lbl_rev.configure(text='No se pudo completar la revisión (ver Avisos).'); self.nb.tab(2, text='Revisión')
            for w in (self.btn_aceptar, self.btn_quitar, self.btn_rev_xls, self.btn_hoja): w.state(['disabled'])
            return
        for i, x in d.iterrows():
            ok = str(x.Estado).startswith('Aceptada')
            self.tv_rev.insert('', 'end', iid=str(i), values=('●', x.Control, f'{x.Servicio} {str(x.Programa).split(" - ")[0]}'.strip(), x.Detalle,
                                                              x.Estado), tags=('ok',) if ok else (x.Nivel,))
        rs = RV.resumen(d)
        self.lbl_rev.configure(text=f'{rs["rojos"]} rojo(s) · {rs["amarillos"]} amarillo(s) · {rs["aceptadas"]} aceptada(s) con comentario. '
                                    + ('Quedan rojos sin explicar: corríjalos en el origen y vuelva a generar, o acéptelos con un comentario.' if rs['rojos_pendientes']
                                       else 'No quedan rojos sin explicar: puede exportar la hoja de control.') + '  Doble clic para ver el detalle.')
        self.nb.tab(2, text=f'Revisión ({rs["rojos_pendientes"]} rojos)' if rs['rojos_pendientes'] else 'Revisión ✔')
        for w in (self.btn_aceptar, self.btn_quitar, self.btn_rev_xls): w.state(['!disabled'])
        self.btn_hoja.state(['disabled'] if rs['rojos_pendientes'] else ['!disabled'])
        self.btn_hoja.configure(style='Secundario.TButton' if rs['rojos_pendientes'] else 'Primario.TButton')

    def _seleccion(self):
        return [int(i) for i in self.tv_rev.selection()]

    def ver_alerta(self):
        sel = self._seleccion()
        if not sel: return
        x = self.resultado['revision'].loc[sel[0]]
        txt = f'{x.Nivel} · {x.Control}\n{x.Servicio} {x.Programa}\n{x["Código"]} {x.Nombre}\n\n{x.Detalle}\n\nEstado: {x.Estado}'
        if x.Comentario: txt += f'\nComentario: {x.Comentario} ({x.Usuario}, {x.Fecha})'
        messagebox.showinfo('Detalle de la alerta', txt)

    def aceptar_alertas(self):
        sel = self._seleccion()
        if not sel:
            messagebox.showinfo('Elija alertas', 'Seleccione una o más alertas de la lista (Ctrl+clic para varias).'); return
        c = simpledialog.askstring('Aceptar con comentario', f'Explique por qué se acepta{"n" if len(sel) > 1 else ""} '
                                   f'{"estas " + str(len(sel)) + " alertas" if len(sel) > 1 else "esta alerta"} '
                                   '(por ejemplo: "decreto N°xxx en trámite" o "excedible por ley"):', parent=self)
        if c is None: return
        if not c.strip():
            messagebox.showwarning('Falta el comentario', 'Para aceptar una alerta hay que escribir un comentario.'); return
        d = self.resultado['revision']
        for i in sel: RV.aceptar(d, i, c)
        self._bitacora('Alerta aceptada', [d.loc[i].to_dict() for i in sel])
        self._guardar_revision()

    def quitar_aceptacion(self):
        sel = self._seleccion()
        if not sel: return
        d = self.resultado['revision']
        for i in sel: RV.quitar(d, i)
        self._bitacora('Aceptación quitada', [d.loc[i].to_dict() for i in sel])
        self._guardar_revision()

    def _bitacora(self, accion, filas, comentario=''):
        r = self.resultado
        try:
            import bitacora as BT
            BT.registrar_revision(self.cfg['carpeta_informes'], r['anio'], r['mes'], accion, filas, comentario)
        except Exception as e:
            self.txt_log.insert('end', f'No se pudo anotar en la bitácora: {e}\n')

    def _guardar_revision(self):
        r = self.resultado
        try:
            RV.guardar(r['revision'], r['carpeta'], r['anio'], r['mes'], r['corte'])
        except PermissionError as e:
            messagebox.showerror('No se pudo guardar', str(e))
        self.mostrar_revision()

    def exportar_hoja(self):
        r = self.resultado
        try:
            ruta = RV.hoja_de_control(r['revision'], r['ctrl'], r['carpeta'], r['anio'], r['mes'], r['corte'], r['salida'], VERSION)
        except (ValueError, PermissionError) as e:
            messagebox.showwarning('Hoja de control', str(e)); return
        self._bitacora('Hoja de control exportada', None, os.path.basename(ruta))
        if messagebox.askyesno('Hoja de control', f'Se guardó {os.path.basename(ruta)} en la carpeta MINAGRI del mes.\n\n¿Abrirla ahora?'):
            abrir(ruta)

    def agregar_al_catalogo(self):
        if self.resultado and self.resultado.get('pendientes'):
            VentanaCatalogo(self, self.resultado['pendientes'])

    def resumen_ejecutivo(self):
        """Se vuelve a preparar al abrirlo, para que incluya las explicaciones de la revisión registradas después de generar."""
        if not self.resultado: return
        try:
            import documentos as DOC
            abrir(DOC.resumen_ejecutivo(self.resultado))
        except Exception as e:
            messagebox.showerror('Resumen ejecutivo', str(e))

    def ver_ideas(self):
        if self.resultado: VerIdeas(self, self.resultado)


class Configuracion(tk.Toplevel):
    CAMPOS = (('carpeta_informes', 'Carpeta de informes mensuales', 'Contiene las carpetas "01 Enero 2026", "02 Febrero 2026", …', 'dir'),
              ('carpeta_presupuesto_vigente', 'Carpeta de Presupuesto Vigente', 'Contiene las 6 planillas "0X SERVICIO Presupuesto Vigente 2026.xlsx"', 'dir'),
              ('carpeta_informes_anio_anterior', 'Carpeta de informes del año anterior (opcional)', 'Para comparar las Ideas Fuerza con igual mes del año anterior', 'dir'),
              ('catalogo', 'Catálogo Maestro', 'Puede ser el catálogo compartido de la unidad (SharePoint), para que todos usen el mismo; vacío = el que trae el agente', 'file'),
              ('carpeta_actualizaciones', 'Carpeta compartida de actualizaciones (opcional)', 'Donde se dejan los .zip de versiones nuevas; la ventana avisa al abrirse si hay una', 'dir'),
              ('carpeta_decretos', 'Carpeta con PDF de decretos (opcional)', 'Se busca aquí el decreto antes de descargarlo de DIPRES (nombre del PDF con el N° del decreto)', 'dir'))

    def __init__(self, app):
        super().__init__(app); self.app = app
        self.title('Configurar carpetas'); self.configure(bg=BLANCO, padx=22, pady=18); self.transient(app); self.grab_set(); self.resizable(True, False)
        try: cfg = leer_config()
        except Exception: cfg = {k: '' for k, *_ in self.CAMPOS}
        ttk.Label(self, text='Carpetas que usa el agente', style='Paso.TLabel').pack(anchor='w')
        ttk.Label(self, text='Se configuran una sola vez. Use "Examinar…" para elegir cada carpeta.', style='Suave.TLabel').pack(anchor='w', pady=(0, 10))
        self.vars = {}
        for k, t, ayuda, tipo in self.CAMPOS:
            ttk.Label(self, text=t, font=app.f['neg']).pack(anchor='w', pady=(8, 0))
            ttk.Label(self, text=ayuda, style='Suave.TLabel').pack(anchor='w')
            f = tk.Frame(self, bg=BLANCO); f.pack(fill='x', pady=(4, 0))
            v = tk.StringVar(value=cfg.get(k, '')); self.vars[k] = v
            ttk.Entry(f, textvariable=v, width=90).pack(side='left', fill='x', expand=True)
            ttk.Button(f, text='Examinar…', style='Secundario.TButton', command=lambda v=v, tipo=tipo: self.examinar(v, tipo)).pack(side='left', padx=(8, 0))
        b = tk.Frame(self, bg=BLANCO); b.pack(fill='x', pady=(18, 0))
        ttk.Button(b, text='Guardar', style='Primario.TButton', command=self.guardar).pack(side='right')
        ttk.Button(b, text='Cancelar', style='Secundario.TButton', command=self.destroy).pack(side='right', padx=8)
        ttk.Button(b, text='Preparar año nuevo…', style='Secundario.TButton', command=lambda: (self.destroy(), AnioNuevo(self.app))).pack(side='left')

    def examinar(self, v, tipo):
        ini = v.get() if os.path.exists(v.get()) else os.path.expanduser('~')
        r = filedialog.askdirectory(parent=self, initialdir=ini) if tipo == 'dir' else filedialog.askopenfilename(parent=self, initialdir=os.path.dirname(ini) if os.path.isfile(ini) else ini, filetypes=[('Excel', '*.xlsx')])
        if r: v.set(os.path.normpath(r))

    def guardar(self):
        cfg = {k: v.get().strip() for k, v in self.vars.items()}
        if not cfg['catalogo']: cfg['catalogo'] = leer_config()['catalogo'] if os.path.exists(CONFIG) else CATALOGO_POR_DEFECTO
        prob = revisar_rutas(cfg)
        if prob and not messagebox.askyesno('Revisar carpetas', '• ' + '\n• '.join(prob) + '\n\n¿Guardar de todos modos?', parent=self): return
        try: guardar_config(cfg)
        except Exception as e:
            messagebox.showerror('No se pudo guardar', str(e), parent=self); return
        self.destroy(); self.app.cargar_config()


class AnioNuevo(tk.Toplevel):
    """Asistente de cambio de año: catálogo del año nuevo, carpetas y año anterior."""
    def __init__(self, app):
        super().__init__(app); self.app = app
        self.title('Preparar año nuevo'); self.configure(bg=BLANCO, padx=22, pady=18); self.transient(app); self.grab_set()
        cfg = leer_config(); self.cfg = cfg
        import re
        m = re.search(r'(\d{4})\.xlsx$', cfg['catalogo']); actual = int(m.group(1)) if m else datetime.date.today().year
        ttk.Label(self, text='Preparar el agente para el año nuevo', style='Paso.TLabel').pack(anchor='w')
        ttk.Label(self, text=f'Se crea el Catálogo Maestro del año nuevo a partir del de {actual} (sin fechas de corte), la carpeta de informes de {actual} '
                             'pasa a ser la del año anterior y se apuntan las carpetas nuevas. No se modifica nada del año que termina.',
                  style='Suave.TLabel', wraplength=640, justify='left').pack(anchor='w', pady=(0, 10))
        self.v_anio = tk.StringVar(value=str(actual + 1)); self.v_inf = tk.StringVar(); self.v_pv = tk.StringVar()
        for t, v, ayuda in (('Año nuevo', self.v_anio, ''), ('Carpeta de informes mensuales del año nuevo', self.v_inf, f'La que tendrá "01 Enero {actual + 1}", "02 Febrero {actual + 1}", …'),
                            ('Carpeta de Presupuesto Vigente del año nuevo', self.v_pv, 'Puede quedar vacía si aún no están las planillas; se configura después')):
            ttk.Label(self, text=t, font=app.f['neg']).pack(anchor='w', pady=(8, 0))
            if ayuda: ttk.Label(self, text=ayuda, style='Suave.TLabel').pack(anchor='w')
            f = tk.Frame(self, bg=BLANCO); f.pack(fill='x', pady=(4, 0))
            ttk.Entry(f, textvariable=v, width=80 if v is not self.v_anio else 8).pack(side='left', fill='x', expand=v is not self.v_anio)
            if v is not self.v_anio:
                ttk.Button(f, text='Examinar…', style='Secundario.TButton',
                           command=lambda v=v: self.elegir(v)).pack(side='left', padx=(8, 0))
        self.txt = tk.Text(self, height=9, wrap='word', relief='flat', bg=GRIS_FONDO, font=app.f['chica']); self.txt.pack(fill='both', expand=True, pady=(12, 0))
        b = tk.Frame(self, bg=BLANCO); b.pack(fill='x', pady=(12, 0))
        self.btn = ttk.Button(b, text='Preparar', style='Primario.TButton', command=self.preparar); self.btn.pack(side='right')
        ttk.Button(b, text='Cerrar', style='Secundario.TButton', command=self.destroy).pack(side='right', padx=8)

    def preparar(self):
        try: anio = int(self.v_anio.get())
        except ValueError:
            messagebox.showwarning('Año', 'Escriba el año nuevo con cuatro dígitos.', parent=self); return
        inf = self.v_inf.get().strip()
        if not inf or not os.path.isdir(inf):
            messagebox.showwarning('Carpeta de informes', 'Elija la carpeta de informes mensuales del año nuevo (debe existir).', parent=self); return
        if not messagebox.askyesno('Preparar año nuevo', f'Se creará el catálogo {anio} y se cambiarán las carpetas en config.ini (queda un respaldo).\n\n¿Continuar?', parent=self): return
        try:
            import anio_nuevo as AN
            lineas = AN.preparar(self.cfg, anio, inf, self.v_pv.get().strip())
        except Exception as e:
            messagebox.showerror('No se pudo preparar', str(e), parent=self); return
        self.txt.delete('1.0', 'end'); self.txt.insert('1.0', '\n'.join(lineas)); self.btn.state(['disabled'])
        self.app.cargar_config()

    def elegir(self, v):
        r = filedialog.askdirectory(parent=self)
        if r: v.set(os.path.normpath(r))


class Actualizaciones(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app); self.app = app
        self.title('Actualizaciones del agente'); self.configure(bg=BLANCO, padx=22, pady=18); self.transient(app); self.grab_set()
        self.geometry('760x560'); self.minsize(700, 520)
        v = ACT.leer_version()
        ttk.Label(self, text=f'Versión instalada: {v.get("version", VERSION)}', style='Paso.TLabel').pack(anchor='w')
        ttk.Label(self, text=f'Fecha: {v.get("fecha", "") or "sin registro"}', style='Suave.TLabel').pack(anchor='w')
        if v.get('cambios'):
            ttk.Label(self, text='Novedades de esta versión:', font=app.f['neg']).pack(anchor='w', pady=(10, 2))
            fn = tk.Frame(self, bg=BLANCO); fn.pack(fill='x')
            sbn = ttk.Scrollbar(fn, orient='vertical')
            tn = tk.Text(fn, wrap='word', height=8, relief='flat', bg=GRIS_FONDO, fg=GRIS_TEXTO, font=app.f['normal'], padx=8, pady=6,
                         yscrollcommand=sbn.set, cursor='arrow', highlightthickness=0, spacing3=3)
            sbn.configure(command=tn.yview); sbn.pack(side='right', fill='y'); tn.pack(side='left', fill='x', expand=True)
            tn.tag_configure('item', lmargin1=4, lmargin2=22)
            for k, c in CRIT_COLOR.items(): tn.tag_configure('c_' + k, foreground=c, font=app.f['neg'])
            for c in v['cambios']:
                k, txt = separar_criticidad(c)
                tn.insert('end', '●  ', ('item', 'c_' + k) if k else 'item'); tn.insert('end', (f'{k}: ' if k else '') + txt + '\n', 'item')
            tn.configure(state='disabled')

        ttk.Label(self, text='Instalar una versión nueva', font=app.f['neg']).pack(anchor='w', pady=(16, 2))
        ttk.Label(self, text='Elija el .zip que recibió. Su configuración y el Catálogo Maestro se conservan, y la versión actual se respalda.',
                  style='Suave.TLabel', wraplength=700, justify='left').pack(anchor='w')
        f = tk.Frame(self, bg=BLANCO); f.pack(fill='x', pady=(6, 0))
        ttk.Button(f, text='Instalar desde un archivo .zip…', style='Primario.TButton', command=self.desde_archivo).pack(side='left')
        carpeta = app.cfg.get('carpeta_actualizaciones', '') if hasattr(app, 'cfg') else ''
        if carpeta:
            ttk.Button(f, text='Buscar en la carpeta compartida', style='Secundario.TButton', command=lambda: self.buscar(carpeta)).pack(side='left', padx=8)
        else:
            pass
        repo = ACT.repositorio_github(app.cfg.get('repositorio_github', '') if hasattr(app, 'cfg') else '')
        if repo:
            ttk.Button(f, text='Buscar en GitHub', style='Secundario.TButton', command=lambda: self.buscar_github(repo)).pack(side='left', padx=8)
        if not carpeta and not repo:
            ttk.Label(f, text='Puede configurar una carpeta compartida en "Configurar carpetas" para recibir avisos.', style='Suave.TLabel').pack(side='left', padx=10)

        ttk.Label(self, text='Volver a una versión anterior', font=app.f['neg']).pack(anchor='w', pady=(18, 2))
        self.resp = ACT.respaldos()
        lf = tk.Frame(self, bg=BLANCO); lf.pack(fill='both', expand=True, pady=(4, 0))
        self.lista = tk.Listbox(lf, height=6, font=app.f['normal'], activestyle='none', selectbackground=AZUL, borderwidth=1, relief='solid', highlightthickness=0)
        for p, ver, fecha in self.resp: self.lista.insert('end', f'  Versión {ver}   ·   respaldada el {fecha}')
        if not self.resp: self.lista.insert('end', '  (Aún no hay respaldos: se crean al instalar una actualización)')
        self.lista.pack(fill='both', expand=True)
        b = tk.Frame(self, bg=BLANCO); b.pack(fill='x', pady=(10, 0))
        ttk.Button(b, text='Volver a la versión seleccionada', style='Secundario.TButton', command=self.volver).pack(side='left')
        ttk.Button(b, text='Ver historial de versiones', style='Secundario.TButton', command=lambda: HistorialVersiones(app, parent=self)).pack(side='left', padx=8)
        ttk.Button(b, text='Cerrar', style='Secundario.TButton', command=self.destroy).pack(side='right')

    def desde_archivo(self):
        r = filedialog.askopenfilename(parent=self, title='Elegir el .zip de la actualización', filetypes=[('Actualización del agente', '*.zip')])
        if not r: return
        try: info = ACT.inspeccionar(r)
        except Exception as e:
            messagebox.showerror('Archivo no válido', str(e), parent=self); return
        self.confirmar(info)

    def buscar(self, carpeta):
        info = ACT.buscar_en_carpeta(carpeta, VERSION)
        if not info:
            messagebox.showinfo('Sin novedades', f'No hay versiones más nuevas que la {VERSION} en:\n{carpeta}', parent=self); return
        self.confirmar(info)

    def buscar_github(self, repo):
        try:
            self.configure(cursor='watch'); self.update()
            info = ACT.buscar_en_github(repo, VERSION)
        except Exception as e:
            self.configure(cursor='')
            messagebox.showwarning('Sin conexión con GitHub', f'No se pudo consultar github.com/{repo}.\n\n{e}\n\n'
                                   'Revise la conexión a Internet o pida el .zip y use "Instalar desde un archivo .zip…".', parent=self); return
        self.configure(cursor='')
        if not info:
            messagebox.showinfo('Sin novedades', f'Tiene la versión más reciente publicada en GitHub ({VERSION}).', parent=self); return
        self.confirmar(info)

    def confirmar(self, info):
        aviso = '' if ACT.es_mas_nueva(info['version'], VERSION) else f'\n\nATENCIÓN: no es más nueva que la instalada ({VERSION}).'
        cambios = '\n'.join('• ' + c for c in info.get('cambios', [])[:8])
        if messagebox.askyesno('Instalar actualización', f'Versión {info["version"]} ({info.get("fecha", "")}).\n\n{cambios}{aviso}\n\n¿Instalar ahora?', parent=self):
            self.app.instalar_actualizacion(info, ventana=self)

    def volver(self):
        sel = self.lista.curselection()
        if not self.resp or not sel:
            messagebox.showinfo('Elija una versión', 'Seleccione en la lista la versión a la que quiere volver.', parent=self); return
        p, ver, fecha = self.resp[sel[0]]
        if not messagebox.askyesno('Volver a una versión anterior', f'¿Volver a la versión {ver} (respaldada el {fecha})?\n\n'
                                   'La versión actual también se respalda, así que puede deshacer este paso.', parent=self): return
        try: ACT.restaurar(p, log=lambda t: None)
        except Exception as e:
            messagebox.showerror('No se pudo volver', str(e), parent=self); return
        messagebox.showinfo('Listo', f'Se volvió a la versión {ver}. La ventana se abrirá de nuevo.', parent=self)
        self.app.reiniciar()


class HistorialVersiones(tk.Toplevel):
    """Resumen de los cambios de cada versión (programa/historial.json), de la más nueva a la más antigua."""
    def __init__(self, app, parent=None):
        super().__init__(parent or app); self.app = app
        self.title('Historial de versiones'); self.configure(bg=BLANCO); self.transient(parent or app)
        self.geometry('640x520'); self.minsize(520, 380)
        try:
            with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'historial.json'), encoding='utf-8') as f:
                versiones = json.load(f)
        except Exception:
            versiones = []
        top = tk.Frame(self, bg=BLANCO, padx=20, pady=14); top.pack(fill='x')
        ttk.Label(top, text='Historial de versiones', style='Paso.TLabel').pack(anchor='w')
        ttk.Label(top, text=f'Versión instalada: {VERSION}', style='Suave.TLabel').pack(anchor='w')
        b = tk.Frame(self, bg=BLANCO, padx=20, pady=12); b.pack(fill='x', side='bottom')
        cuerpo = tk.Frame(self, bg=BLANCO, padx=20); cuerpo.pack(fill='both', expand=True)
        barra = ttk.Scrollbar(cuerpo, orient='vertical')
        t = tk.Text(cuerpo, wrap='word', relief='flat', bg=GRIS_FONDO, fg=GRIS_TEXTO, font=app.f['normal'], padx=16, pady=12,
                    spacing1=2, spacing3=4, yscrollcommand=barra.set, cursor='arrow', highlightthickness=0)
        barra.configure(command=t.yview); barra.pack(side='right', fill='y'); t.pack(side='left', fill='both', expand=True)
        t.tag_configure('ver', font=app.f['paso'], foreground=AZUL_OSC, spacing1=10)
        t.tag_configure('meta', font=app.f['chica'], foreground=GRIS_SUAVE)
        t.tag_configure('actual', font=app.f['chica'], foreground=BLANCO, background=VERDE)
        t.tag_configure('res', font=app.f['neg'], foreground=GRIS_TEXTO)
        t.tag_configure('item', lmargin1=14, lmargin2=28)
        for k, c in CRIT_COLOR.items():
            t.tag_configure('c_' + k, foreground=c, font=app.f['neg'])
        t.insert('end', 'Criticidad: ', 'meta')
        for k, d in (('Crítica', 'cambia o protege las cifras'), ('Media', 'controles y procesos'), ('Baja', 'ventana y documentos')):
            t.insert('end', '● ', 'c_' + k); t.insert('end', f'{k} ({d})   ', 'meta')
        t.insert('end', '\n')
        if not versiones:
            t.insert('end', 'No se encontró el resumen de versiones.')
        for v in versiones:
            t.insert('end', f'Versión {v["version"]}' + (' beta' if ACT.es_beta(v['version']) else ''), 'ver')
            if str(v['version']) == str(VERSION): t.insert('end', '   (instalada)', 'meta')
            t.insert('end', '\n')
            if v.get('criticidad'): t.insert('end', '● ', 'c_' + v['criticidad']); t.insert('end', f'Criticidad {v["criticidad"].lower()} · ', 'meta')
            t.insert('end', f'{v.get("tipo", "")} · {v.get("fecha", "")}\n', 'meta')
            t.insert('end', v.get('resumen', '') + '\n', 'res')
            for d in v.get('cambios', []):
                t.insert('end', '●  ', ('item', 'c_' + d.get('criticidad', 'Media'))); t.insert('end', f'{d.get("texto", "")}\n', 'item')
            for d in v.get('detalle', []): t.insert('end', f'•  {d}\n', 'item')
        t.configure(state='disabled')
        if os.path.exists(HISTORIAL):
            ttk.Button(b, text='Abrir el historial completo (PDF)', style='Secundario.TButton', command=lambda: abrir(HISTORIAL)).pack(side='left')
        ttk.Button(b, text='Cerrar', style='Secundario.TButton', command=self.cerrar).pack(side='right')
        self.bind('<Escape>', lambda e: self.cerrar())
        self.protocol('WM_DELETE_WINDOW', self.cerrar)
        # Si se abre desde una ventana que tiene la captura de eventos (Actualizaciones), en macOS esa ventana
        # sigue recibiendo los clics y los botones de esta no responden: la captura pasa a esta ventana y se
        # devuelve al cerrar.
        self._grab_anterior = None
        try:
            g = self.grab_current()
            self._grab_anterior = g if g is not None and g is not self else None
            self.lift(); self.focus_force(); self.after(50, self._tomar_grab)
        except tk.TclError:
            pass

    def _tomar_grab(self):
        try: self.grab_set()
        except tk.TclError: pass

    def cerrar(self):
        try: self.grab_release()
        except tk.TclError: pass
        anterior = self._grab_anterior
        self.destroy()
        try:
            if anterior is not None and anterior.winfo_exists(): anterior.grab_set(); anterior.focus_force()
        except tk.TclError:
            pass


class VerIdeas(tk.Toplevel):
    def __init__(self, app, r):
        super().__init__(app); self.title(f'Ideas Fuerza – borrador {r["nombre"]}'); self.geometry('900x640'); self.configure(bg=BLANCO)
        ideas = r['ideas']
        top = tk.Frame(self, bg=BLANCO, padx=16, pady=12); top.pack(fill='x')
        ttk.Label(top, text='Borrador de Ideas Fuerza', style='Paso.TLabel').pack(anchor='w')
        ttk.Label(top, text='Generado con las cifras del mes. Complete los textos entre corchetes [Completar: …] antes de publicar.', style='Suave.TLabel').pack(anchor='w')
        cuerpo = tk.Frame(self, bg=BLANCO, padx=16); cuerpo.pack(fill='both', expand=True)
        lista = tk.Listbox(cuerpo, width=16, font=app.f['neg'], activestyle='none', relief='flat', highlightthickness=1, highlightbackground=GRIS_BORDE,
                           selectbackground=AZUL, selectforeground=BLANCO)
        lista.pack(side='left', fill='y')
        for n in ideas['Institución']: lista.insert('end', n)
        self.txt = tk.Text(cuerpo, wrap='word', font=app.f['normal'], relief='flat', padx=14, pady=10, bg=GRIS_FONDO, fg=GRIS_TEXTO, spacing2=2, spacing3=8)
        self.txt.pack(side='left', fill='both', expand=True, padx=(12, 0))
        self.txt.tag_configure('marca', background='#fff3cd', foreground='#7a5b00')
        def mostrar(e=None):
            i = lista.curselection()[0] if lista.curselection() else 0
            self.txt.delete('1.0', 'end'); self.txt.insert('1.0', ideas['Ideas Fuerza'].iat[i])
            ini = '1.0'
            while True:
                p = self.txt.search('[Completar', ini, stopindex='end')
                if not p: break
                fin = self.txt.search(']', p, stopindex='end') or 'end'
                self.txt.tag_add('marca', p, f'{fin}+1c'); ini = f'{fin}+1c'
        lista.bind('<<ListboxSelect>>', mostrar); lista.selection_set(0); mostrar()
        b = tk.Frame(self, bg=BLANCO, padx=16, pady=12); b.pack(fill='x', side='bottom', before=cuerpo)
        ttk.Button(b, text='Copiar texto', style='Primario.TButton', command=self.copiar).pack(side='left')
        ttk.Button(b, text='Abrir en Word', style='Secundario.TButton', command=lambda: self.word(r)).pack(side='left', padx=8)
        ttk.Button(b, text='Cerrar', style='Secundario.TButton', command=self.destroy).pack(side='right')

    def word(self, r):
        try:
            import documentos as DOC
            abrir(DOC.ideas_fuerza(r))
        except Exception as e:
            messagebox.showerror('Ideas Fuerza en Word', str(e), parent=self)

    def copiar(self):
        self.clipboard_clear(); self.clipboard_append(self.txt.get('1.0', 'end-1c'))
        messagebox.showinfo('Copiado', 'El texto quedó copiado. Péguelo en su minuta o en la hoja Ideas Fuerza.', parent=self)



class VentanaCatalogo(tk.Toplevel):
    """Lo que el mes trae y el Catálogo Maestro no tiene: se completa aquí y se agrega al catálogo."""
    def __init__(self, app, pendientes):
        super().__init__(app); self.app = app; self.title('Agregar al Catálogo Maestro'); self.configure(bg=BLANCO)
        self.transient(app); self.geometry('880x520')
        import editor_catalogo as EC
        self.EC = EC; self.pend = pendientes
        try: self.op = EC.opciones(app.cfg['catalogo'])
        except Exception as e:
            messagebox.showerror('Catálogo', f'No se pudo leer el Catálogo Maestro: {e}', parent=self); self.destroy(); return
        ttk.Label(self, text='Complete lo que corresponda y marque lo que quiere agregar. Solo se agregan filas nuevas; '
                             'el catálogo queda respaldado antes de guardar.', style='Suave.TLabel', wraplength=840).pack(anchor='w', padx=16, pady=(12, 6))
        marco = tk.Frame(self, bg=BLANCO); marco.pack(fill='both', expand=True, padx=16)
        cv = tk.Canvas(marco, bg=BLANCO, highlightthickness=0); sb = ttk.Scrollbar(marco, orient='vertical', command=cv.yview)
        dentro = tk.Frame(cv, bg=BLANCO); dentro.bind('<Configure>', lambda e: cv.configure(scrollregion=cv.bbox('all')))
        cv.create_window((0, 0), window=dentro, anchor='nw'); cv.configure(yscrollcommand=sb.set)
        cv.pack(side='left', fill='both', expand=True); sb.pack(side='right', fill='y')
        self.filas = []
        for i, p in enumerate(pendientes):
            f = tk.Frame(dentro, bg=BLANCO); f.pack(fill='x', pady=4)
            usar = tk.BooleanVar(value=True)
            ttk.Checkbutton(f, variable=usar).grid(row=0, column=0, rowspan=2, sticky='n')
            tk.Label(f, text=EC.describir(p), bg=BLANCO, fg=GRIS_TEXTO, font=app.f['chica'], wraplength=780, justify='left').grid(row=0, column=1, columnspan=4, sticky='w')
            v = {}
            if p['tipo'] == 'Programa':
                v['tratamiento'] = tk.StringVar(value='Incluir'); v['programa'] = tk.StringVar()
                ttk.Label(f, text='Tratamiento', style='Suave.TLabel').grid(row=1, column=1, sticky='w')
                ttk.Combobox(f, textvariable=v['tratamiento'], values=EC.TRATAMIENTOS, state='readonly', width=12).grid(row=1, column=2, sticky='w', padx=4)
                ttk.Label(f, text='Programa', style='Suave.TLabel').grid(row=1, column=3, sticky='w')
                ttk.Combobox(f, textvariable=v['programa'], values=self.op['programas'].get(p['servicio'], []), width=58).grid(row=1, column=4, sticky='w', padx=4)
            elif p['tipo'] in ('Región', 'Región CF'):
                v['region'] = tk.StringVar()
                ttk.Label(f, text='Región', style='Suave.TLabel').grid(row=1, column=1, sticky='w')
                ttk.Combobox(f, textvariable=v['region'], values=EC.REGIONES, state='readonly', width=24).grid(row=1, column=2, sticky='w', padx=4)
            elif p['tipo'] == 'Nombre':
                v['nombre'] = tk.StringVar(value=p.get('sugerido') or f'{p["valor"]} ')
                ttk.Label(f, text='Nombre', style='Suave.TLabel').grid(row=1, column=1, sticky='w')
                ttk.Entry(f, textvariable=v['nombre'], width=70).grid(row=1, column=2, columnspan=3, sticky='w', padx=4)
            elif p['tipo'] == 'Ingreso':
                v['codigo'] = tk.StringVar()
                ttk.Label(f, text='Código clasificador', style='Suave.TLabel').grid(row=1, column=1, sticky='w')
                ttk.Entry(f, textvariable=v['codigo'], width=12).grid(row=1, column=2, sticky='w', padx=4)
            self.filas.append((p, usar, v))
        b = tk.Frame(self, bg=BLANCO); b.pack(fill='x', padx=16, pady=12)
        ttk.Button(b, text='Agregar al catálogo', style='Primario.TButton', command=self.guardar).pack(side='right')
        ttk.Button(b, text='Cancelar', style='Secundario.TButton', command=self.destroy).pack(side='right', padx=8)

    def guardar(self):
        filas, faltan = [], []
        for p, usar, v in self.filas:
            if not usar.get(): continue
            val = {k: x.get().strip() for k, x in v.items()}
            if (p['tipo'] == 'Programa' and val['tratamiento'] != 'Excluir' and not val['programa']) or \
               (p['tipo'] in ('Región', 'Región CF') and not val['region']) or \
               (p['tipo'] == 'Nombre' and len(val['nombre']) <= len(str(p['valor'])) + 1) or \
               (p['tipo'] == 'Ingreso' and not val['codigo'].isdigit()):
                faltan.append(self.EC.describir(p)); continue
            filas.append(self.EC.fila_para(p, val, self.op))
        if faltan:
            messagebox.showwarning('Faltan datos', 'Complete (o desmarque) estos elementos:\n\n' + '\n'.join(faltan[:10]), parent=self); return
        if not filas: self.destroy(); return
        try:
            resp = self.EC.agregar(self.app.cfg['catalogo'], filas)
        except Exception as e:
            messagebox.showerror('No se pudo guardar', str(e), parent=self); return
        messagebox.showinfo('Catálogo actualizado', f'Se agregaron {len(filas)} fila(s) al Catálogo Maestro.\n'
                            f'Respaldo del catálogo anterior: respaldos/catálogo/{os.path.basename(resp)}\n\n'
                            'Vuelva a generar el mes para aplicar los cambios.', parent=self)
        self.destroy()
