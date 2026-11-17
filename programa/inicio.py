"""Ventana de bienvenida de los Agentes de Apoyo a la Gestión (UPROGE, Ministerio de Agricultura).

Muestra un botón por cada agente disponible. Para agregar un agente nuevo (por ejemplo, el informe semanal),
basta con sumar una entrada a AGENTES con su título, una línea de descripción y la función que abre su ventana."""
import os, sys, json, threading, subprocess, configparser
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import tkinter as tk
from tkinter import ttk, messagebox, font as tkfont
import actualizador as ACT

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RECURSOS = os.path.join(RAIZ, 'recursos')
AZUL = '#006fb3'; AZUL_HOVER = '#005a92'; AZUL_OSC = '#0a132d'; ROJO = '#fe6565'
GRIS_FONDO = '#f8f9fa'; GRIS_BORDE = '#dee2e6'; GRIS_SUAVE = '#6c757d'; BLANCO = '#ffffff'


def _abrir_ejecucion():
    import gui_ejecucion
    app = gui_ejecucion.App()
    app.mainloop()
    return getattr(app, 'volver_al_inicio', False)

# Agentes disponibles, en el orden en que aparecen los botones.
AGENTES = [
    {'titulo': 'Ejecución Presupuestaria - Power BI',
     'detalle': 'Consolidado mensual MINAGRI para el Informe de Ejecución en Power BI',
     'abrir': _abrir_ejecucion},
]


def version():
    try:
        with open(os.path.join(RAIZ, 'version.json'), encoding='utf-8') as f:
            return json.load(f).get('version', '')
    except Exception:
        return ''


class Bienvenida(tk.Tk):
    def __init__(self):
        super().__init__()
        self.eleccion = None
        self.title('Agentes de Apoyo a la Gestión – UPROGE')
        self.configure(bg=BLANCO); self.resizable(False, False)
        self._fuentes(); self._icono()

        # franja institucional
        franja = tk.Frame(self, height=6, bg=AZUL); franja.pack(fill='x')
        tk.Frame(franja, height=6, width=200, bg=ROJO).pack(side='left')
        # aviso de actualización (se muestra solo si hay una versión nueva)
        self.aviso = tk.Frame(self, bg='#e6f1f8', padx=16, pady=8)
        self.aviso_txt = tk.Label(self.aviso, font=self.f['chica'], bg='#e6f1f8', fg=AZUL_OSC, justify='left')
        self.aviso_txt.pack(side='left')
        tk.Button(self.aviso, text='Ver e instalar', font=self.f['chica'] + ('bold',), bg=AZUL, fg=BLANCO, activebackground=AZUL_HOVER,
                  activeforeground=BLANCO, relief='flat', bd=0, padx=12, pady=4, cursor='hand2', command=self.instalar).pack(side='right')
        self._franja = franja; self.nueva = None

        cuerpo = tk.Frame(self, bg=BLANCO, padx=56, pady=34); cuerpo.pack(fill='both', expand=True)
        self._logo = None
        for n in ('logo_grande.png', 'logo.png'):
            p = os.path.join(RECURSOS, n)
            if os.path.exists(p):
                try:
                    img = tk.PhotoImage(file=p)
                    k = max(1, int(img.height() / 190 + 0.999)); self._logo = img.subsample(k, k) if k > 1 else img
                    break
                except tk.TclError:
                    self._logo = None
        if self._logo is not None:
            tk.Label(cuerpo, image=self._logo, bg=BLANCO).pack(pady=(0, 22))

        tk.Label(cuerpo, text='Ministerio de Agricultura - UPROGE', font=self.f['titulo'], bg=BLANCO, fg=AZUL_OSC).pack()
        tk.Label(cuerpo, text='Agentes de Apoyo a la Gestión', font=self.f['sub'], bg=BLANCO, fg=GRIS_SUAVE).pack(pady=(4, 0))
        tk.Frame(cuerpo, height=1, width=420, bg=GRIS_BORDE).pack(pady=24)

        for ag in AGENTES:
            b = tk.Button(cuerpo, text=ag['titulo'], font=self.f['boton'], bg=AZUL, fg=BLANCO, activebackground=AZUL_HOVER,
                          activeforeground=BLANCO, relief='flat', bd=0, cursor='hand2', padx=28, pady=14, width=34,
                          command=lambda ag=ag: self.elegir(ag))
            b.pack(pady=(0, 4))
            b.bind('<Enter>', lambda e, b=b: b.configure(bg=AZUL_HOVER)); b.bind('<Leave>', lambda e, b=b: b.configure(bg=AZUL))
            tk.Label(cuerpo, text=ag['detalle'], font=self.f['chica'], bg=BLANCO, fg=GRIS_SUAVE).pack(pady=(0, 14))

        pie = tk.Frame(self, bg=GRIS_FONDO, padx=20, pady=8); pie.pack(fill='x', side='bottom')
        v = version()
        tk.Label(pie, text=(f'Versión {v}' if v else '') + ' · Unidad de Programación y Gestión Presupuestaria',
                 font=self.f['chica'], bg=GRIS_FONDO, fg=GRIS_SUAVE).pack()

        self.update_idletasks()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        self.geometry(f'+{max(0, (self.winfo_screenwidth() - w) // 2)}+{max(0, (self.winfo_screenheight() - h) // 3)}')
        if AGENTES: self.bind('<Return>', lambda e: self.elegir(AGENTES[0]))
        if not ACT.REVISADO[0]: threading.Thread(target=self._buscar_version, daemon=True).start()
        elif ACT.ENCONTRADA[0]: self._mostrar_aviso(ACT.ENCONTRADA[0])

    # ---------------------------------------------------------------- actualizaciones
    def _buscar_version(self):
        """Busca una versión nueva en la carpeta compartida y luego en GitHub; sin conexión no avisa nada."""
        c = configparser.ConfigParser()
        try: c.read(os.path.join(RAIZ, 'config.ini'), encoding='utf-8'); g = c['RUTAS']
        except Exception: g = {}
        actual = version() or '0'; info = None
        try: info = ACT.buscar_en_carpeta(g.get('carpeta_actualizaciones', '').strip(), actual)
        except Exception: info = None
        if info is None:
            try: info = ACT.buscar_en_github(ACT.repositorio_github(g.get('repositorio_github', '')), actual)
            except Exception: info = None
        ACT.REVISADO[0] = True
        if info:
            try: self.after(0, lambda: self._mostrar_aviso(info))
            except Exception: pass

    def _mostrar_aviso(self, info):
        if not self.winfo_exists(): return
        self.nueva = info; ACT.YA_OFRECIDAS.add(info['version']); ACT.ENCONTRADA[0] = info
        origen = 'GitHub' if info.get('url') else 'la carpeta compartida'
        self.aviso_txt.configure(text=f'Hay una versión nueva: {info["version"]}' + (f' ({info["fecha"]})' if info.get('fecha') else '')
                                 + f'\nUsted tiene la {version()} · disponible en {origen}')
        self.aviso.pack(fill='x', after=self._franja)

    def instalar(self):
        info = self.nueva
        if not info: return
        cambios = '\n'.join('• ' + c for c in info.get('cambios', [])[:8])
        if not messagebox.askyesno('Instalar versión nueva', f'Versión {info["version"]}' + (f' ({info["fecha"]})' if info.get('fecha') else '')
                                   + f'.\n\n{cambios}\n\n¿Instalarla ahora? Su configuración y el Catálogo Maestro se conservan, '
                                   'y la versión actual queda respaldada.', parent=self): return
        try:
            self.configure(cursor='watch'); self.update()
            ruta = ACT.descargar(info) if info.get('url') else info['archivo']
            r = ACT.instalar(ruta, log=lambda t: None)
        except Exception as e:
            self.configure(cursor='')
            messagebox.showerror('No se pudo actualizar', f'{e}\n\nNo se modificó config.ini ni el Catálogo Maestro.', parent=self); return
        self.configure(cursor='')
        messagebox.showinfo('Actualización instalada', f'Versión {r["version"]} instalada (antes: {r["anterior"]}).\n\n'
                            + '\n'.join('• ' + h for h in r['hechos']) + '\n\nEl programa se abrirá de nuevo con la versión nueva.', parent=self)
        reabrir(); self.destroy()

    def _fuentes(self):
        fam = set(tkfont.families())
        base = next((f for f in ('Roboto', 'Segoe UI', 'Arial', 'Helvetica') if f in fam), 'TkDefaultFont')
        tit = next((f for f in ('Roboto Slab', base) if f in fam), base)
        self.f = {'titulo': (tit, 20, 'bold'), 'sub': (base, 13), 'boton': (base, 12, 'bold'), 'chica': (base, 9)}

    def _icono(self):
        p = os.path.join(RECURSOS, 'agente.png')
        if os.path.exists(p):
            try: self._ico = tk.PhotoImage(file=p); self.iconphoto(True, self._ico)
            except tk.TclError: pass
        ico = os.path.join(RECURSOS, 'agente.ico')
        if sys.platform.startswith('win') and os.path.exists(ico):
            try: self.iconbitmap(ico)
            except tk.TclError: pass

    def elegir(self, ag):
        self.eleccion = ag; self.destroy()


def reabrir():
    ejecutable = sys.executable
    if sys.platform.startswith('win') and ejecutable.lower().endswith('python.exe'):
        w = ejecutable[:-10] + 'pythonw.exe'
        if os.path.exists(w): ejecutable = w
    try:
        subprocess.Popen([ejecutable, os.path.join(RAIZ, 'programa', 'ventana.pyw')], cwd=os.path.join(RAIZ, 'programa'))
    except Exception:
        pass


def main():
    """Bienvenida -> agente elegido -> (si el agente pide volver) bienvenida otra vez."""
    while True:
        b = Bienvenida(); b.mainloop()
        if not b.eleccion: break
        try:
            volver = b.eleccion['abrir']()
        except Exception as e:
            r = tk.Tk(); r.withdraw()
            messagebox.showerror('No se pudo abrir el agente', f'{type(e).__name__}: {e}\n\nSi faltan componentes, ejecute 1_INSTALAR.bat.')
            r.destroy(); volver = True
        if not volver: break


if __name__ == '__main__':
    main()
