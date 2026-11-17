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
RECURSOS = os.path.join(RAIZ_AGENTE, 'recursos')
HISTORIAL = os.path.join(RAIZ_AGENTE, 'Historial de versiones.pdf')


def abrir(path):
    try:
        if sys.platform.startswith('win'): os.startfile(path)
        elif sys.platform == 'darwin': subprocess.Popen(['open', path])
        else: subprocess.Popen(['xdg-open', path])
    except Exception as e:
        messagebox.showerror('No se pudo abrir', f'No se pudo abrir:\n{path}\n\n{e}')


def millones(m): return '$' + f'{m / 1000:,.0f}'.replace(',', '.')
def porc(x): return f'{x * 100:,.1f}'.replace('.', ',') + '%'


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
        self.prog = ttk.Progressbar(t3, mode='determinate', maximum=23, style='Horizontal.TProgressbar'); self.prog.pack(fill='x', pady=(10, 4))
        self.lbl_prog = ttk.Label(t3, text='Listo para comenzar.', style='Suave.TLabel', wraplength=380); self.lbl_prog.pack(anchor='w')

        # Resultados
        tr = tk.Frame(der, bg=BLANCO, highlightbackground=GRIS_BORDE, highlightthickness=1, padx=16, pady=10); tr.pack(fill='both', expand=True, pady=(0, 10))
        h = tk.Frame(tr, bg=BLANCO); h.pack(fill='x')
        ttk.Label(h, text='Resultado', style='Paso.TLabel').pack(side='left')
        self.lbl_res = ttk.Label(h, text='Aún no se ha generado el consolidado.', style='Suave.TLabel'); self.lbl_res.pack(side='left', padx=12)
        self.nb = ttk.Notebook(tr); self.nb.pack(fill='both', expand=True, pady=(8, 8))
        # pestaña ejecución
        p1 = ttk.Frame(self.nb, style='Tarjeta.TFrame', padding=(0, 8)); self.nb.add(p1, text='Ejecución de gastos')
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
        self.btn_quitar = ttk.Button(br, text='Quitar', style='Secundario.TButton', command=self.quitar_aceptacion); self.btn_quitar.pack(side='left', padx=6)
        self.btn_rev_xls = ttk.Button(br, text='Abrir revisión', style='Secundario.TButton', command=lambda: self.resultado and abrir(self.resultado['ruta_revision'])); self.btn_rev_xls.pack(side='left')
        self.btn_hoja = ttk.Button(br, text='Hoja de control', style='Secundario.TButton', command=self.exportar_hoja); self.btn_hoja.pack(side='right')
        for w in (self.btn_aceptar, self.btn_quitar, self.btn_rev_xls, self.btn_hoja): w.state(['disabled'])
        # pestaña validación
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
        # botones de resultado
        b = tk.Frame(tr, bg=BLANCO); b.pack(fill='x')
        self.btn_abrir = ttk.Button(b, text='Abrir Excel generado', style='Primario.TButton', command=lambda: self.resultado and abrir(self.resultado['salida']))
        self.btn_abrir.pack(side='left')
        self.btn_ideas = ttk.Button(b, text='Ver Ideas Fuerza', style='Secundario.TButton', command=self.ver_ideas); self.btn_ideas.pack(side='left', padx=8)
        self.btn_carpeta = ttk.Button(b, text='Abrir carpeta del mes', style='Secundario.TButton', command=lambda: self.estado and abrir(os.path.join(self.estado['carpeta'], 'MINAGRI')))
        self.btn_carpeta.pack(side='left')
        for w in (self.btn_abrir, self.btn_ideas): w.state(['disabled'])

    def _pie(self):
        pie = tk.Frame(self, bg=GRIS_FONDO, padx=20, pady=8); pie.pack(fill='x', side='bottom')
        ttk.Button(pie, text='‹ Inicio', style='Secundario.TButton', width=8, command=self.ir_al_inicio).pack(side='left', padx=(0, 8))
        ttk.Button(pie, text='Configurar carpetas', style='Secundario.TButton', command=self.configurar).pack(side='left')
        ttk.Button(pie, text='Abrir Catálogo Maestro', style='Secundario.TButton', command=lambda: abrir(self.cfg['catalogo'])).pack(side='left', padx=8)
        man = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'Manual de usuario.pdf')
        if os.path.exists(man): ttk.Button(pie, text='Manual de usuario', style='Secundario.TButton', command=lambda: abrir(man)).pack(side='left')
        ttk.Button(pie, text='Actualizaciones', style='Secundario.TButton', command=lambda: Actualizaciones(self)).pack(side='left', padx=8)
        ttk.Button(pie, text='Historial de versiones', style='Secundario.TButton', command=lambda: HistorialVersiones(self)).pack(side='left')
        tk.Label(pie, text=f'Versión {VERSION if ERROR_INICIO is None else "?"}',
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
                try: self.after(0, lambda: self.ofrecer_actualizacion(info))
                except Exception: pass
        threading.Thread(target=buscar, daemon=True).start()

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
            r = ACT.instalar(ruta, log=lambda t: None)
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
        if est['problemas_archivos']:
            self.lbl_oficial.configure(text=est['problemas_archivos'][0])
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
        self.btn_generar.state(['disabled'] if est['problemas_archivos'] else ['!disabled'])

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
        if os.path.exists(self.estado['salida']) and not messagebox.askyesno('Reemplazar', f'Ya existe:\n{os.path.basename(self.estado["salida"])}\n\n¿Desea reemplazarlo?'):
            return
        self.trabajando = True; self.resultado = None
        for w in (self.btn_generar, self.btn_abrir, self.btn_ideas, self.cb_mes, self.btn_sugerir): w.state(['disabled'])
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
        for w in (self.btn_abrir, self.btn_ideas): w.state(['!disabled'])
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
        self.nb.tab(1, text=f'Avisos ({len(rev)} para revisar)' if rev else 'Avisos')
        # validación
        self.tv_val.delete(*self.tv_val.get_children())
        if r['validacion'] is not None:
            for _, x in r['validacion'].iterrows():
                self.tv_val.insert('', 'end', values=(x['Nivel'], f'{int(x["Celdas comparadas"]):,}'.replace(',', '.'), int(x['Con diferencia > 1 (miles de $)'])))
            self.lbl_val.configure(text='Diferencias mayores a 1 (miles de pesos). El detalle de cada una está en la hoja "Validación" del Excel generado.')
        else:
            self.lbl_val.configure(text='No hay consolidado oficial del mes, por lo que no se hizo la comparación.')
        self.mostrar_revision()
        rs = RV.resumen(r['revision']) if r.get('revision') is not None else None
        self.nb.select(2 if rs and rs['rojos_pendientes'] else (1 if rev else 0))
        messagebox.showinfo('Listo', f'Se generó el consolidado de {r["nombre"]}.\n\n{os.path.basename(r["salida"])}\n\n'
                            + (f'Hay {len(rev)} aviso(s) para revisar.' if rev else 'No hay avisos para revisar.')
                            + (f'\n\nRevisión: {rs["rojos_pendientes"]} rojo(s) sin explicar y {rs["amarillos_pendientes"]} amarillo(s).' if rs else ''))

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
        self._guardar_revision()

    def quitar_aceptacion(self):
        sel = self._seleccion()
        if not sel: return
        d = self.resultado['revision']
        for i in sel: RV.quitar(d, i)
        self._guardar_revision()

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
        if messagebox.askyesno('Hoja de control', f'Se guardó {os.path.basename(ruta)} en la carpeta MINAGRI del mes.\n\n¿Abrirla ahora?'):
            abrir(ruta)

    def ver_ideas(self):
        if self.resultado: VerIdeas(self, self.resultado)


class Configuracion(tk.Toplevel):
    CAMPOS = (('carpeta_informes', 'Carpeta de informes mensuales', 'Contiene las carpetas "01 Enero 2026", "02 Febrero 2026", …', 'dir'),
              ('carpeta_presupuesto_vigente', 'Carpeta de Presupuesto Vigente', 'Contiene las 6 planillas "0X SERVICIO Presupuesto Vigente 2026.xlsx"', 'dir'),
              ('carpeta_informes_anio_anterior', 'Carpeta de informes del año anterior (opcional)', 'Para comparar las Ideas Fuerza con igual mes del año anterior', 'dir'),
              ('catalogo', 'Catálogo Maestro', 'Archivo "Catálogo Maestro Ejecución MINAGRI 2026.xlsx"', 'file'),
              ('carpeta_actualizaciones', 'Carpeta compartida de actualizaciones (opcional)', 'Donde se dejan los .zip de versiones nuevas; la ventana avisa al abrirse si hay una', 'dir'))

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

    def examinar(self, v, tipo):
        ini = v.get() if os.path.exists(v.get()) else os.path.expanduser('~')
        r = filedialog.askdirectory(parent=self, initialdir=ini) if tipo == 'dir' else filedialog.askopenfilename(parent=self, initialdir=os.path.dirname(ini) if os.path.isfile(ini) else ini, filetypes=[('Excel', '*.xlsx')])
        if r: v.set(os.path.normpath(r))

    def guardar(self):
        cfg = {k: v.get().strip() for k, v in self.vars.items()}
        if not cfg['catalogo']: cfg['catalogo'] = CATALOGO_POR_DEFECTO
        prob = revisar_rutas(cfg)
        if prob and not messagebox.askyesno('Revisar carpetas', '• ' + '\n• '.join(prob) + '\n\n¿Guardar de todos modos?', parent=self): return
        try: guardar_config(cfg)
        except Exception as e:
            messagebox.showerror('No se pudo guardar', str(e), parent=self); return
        self.destroy(); self.app.cargar_config()


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
            for c in v['cambios']:
                ttk.Label(self, text='•  ' + c, wraplength=700, justify='left').pack(anchor='w')

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
        if not versiones:
            t.insert('end', 'No se encontró el resumen de versiones.')
        for v in versiones:
            t.insert('end', f'Versión {v["version"]}', 'ver')
            if str(v['version']) == str(VERSION): t.insert('end', '   (instalada)', 'meta')
            t.insert('end', '\n'); t.insert('end', f'{v.get("tipo", "")} · {v.get("fecha", "")}\n', 'meta')
            t.insert('end', v.get('resumen', '') + '\n', 'res')
            for d in v.get('detalle', []): t.insert('end', f'•  {d}\n', 'item')
        t.configure(state='disabled')
        if os.path.exists(HISTORIAL):
            ttk.Button(b, text='Abrir el historial completo (PDF)', style='Secundario.TButton', command=lambda: abrir(HISTORIAL)).pack(side='left')
        ttk.Button(b, text='Cerrar', style='Secundario.TButton', command=self.destroy).pack(side='right')
        self.bind('<Escape>', lambda e: self.destroy())


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
        b = tk.Frame(self, bg=BLANCO, padx=16, pady=12); b.pack(fill='x')
        ttk.Button(b, text='Copiar texto', style='Primario.TButton', command=self.copiar).pack(side='left')
        ttk.Button(b, text='Cerrar', style='Secundario.TButton', command=self.destroy).pack(side='right')

    def copiar(self):
        self.clipboard_clear(); self.clipboard_append(self.txt.get('1.0', 'end-1c'))
        messagebox.showinfo('Copiado', 'El texto quedó copiado. Péguelo en su minuta o en la hoja Ideas Fuerza.', parent=self)

