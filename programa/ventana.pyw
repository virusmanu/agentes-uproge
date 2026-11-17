"""Agentes de Apoyo a la Gestión – UPROGE (Ministerio de Agricultura).
Doble clic en 3_ABRIR_AGENTE_VENTANA.bat o en el acceso directo del Escritorio: abre la ventana de bienvenida."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import inicio

if __name__ == '__main__':
    inicio.main()
