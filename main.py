"""
main.py - Punto de arranque.

  python main.py             -> abre la ventana
  python main.py --consola   -> abre el menú de consola (pruebas)

Siempre trabaja con las rutas relativas a ESTA carpeta, sin importar desde
dónde se lance. Así proyecto.db, pdfs_divididos/ y reportes/ no "desaparecen"
ni se duplican si alguien ejecuta el programa desde otro directorio.
"""

import os
import sys

os.chdir(os.path.dirname(os.path.abspath(__file__)))


def iniciar():
    if "--consola" in sys.argv:
        from src import consola

        consola.main()
    else:
        from src.interfaz import iniciar_app

        iniciar_app()


if __name__ == "__main__":
    iniciar()
