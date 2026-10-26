"""
consulta.py - Consultas SQL a proyecto.db desde la terminal (solo lectura).

No forma parte de la aplicación: es una herramienta para revisar la BD.
La conexión es de SOLO LECTURA, así que una consulta no puede dañar los datos
(un UPDATE o DELETE dará error). Se puede usar mientras el programa está abierto.

Uso:
  python consulta.py "SELECT estado, COUNT(*) AS n FROM destinatarios GROUP BY estado"
  python consulta.py                 # modo interactivo: escribe consultas, 'salir' para terminar
  python consulta.py --db otra.db "SELECT * FROM lotes"
"""

import sqlite3
import sys

import pandas as pd

pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 50)
pd.set_option("display.max_colwidth", 60)


def ejecutar(con, sql):
    try:
        df = pd.read_sql_query(sql, con)
    except Exception as e:
        print(f"[ERROR] {e}")
        return
    if df.empty:
        print("(sin resultados)")
    else:
        print(df.to_string(index=False))
        print(f"\n{len(df)} fila(s)")


def main():
    args = sys.argv[1:]
    ruta = "proyecto.db"
    if args[:1] == ["--db"] and len(args) >= 2:
        ruta, args = args[1], args[2:]

    try:
        # mode=ro: solo lectura. timeout: espera si el programa está escribiendo justo ahora.
        con = sqlite3.connect(f"file:{ruta}?mode=ro", uri=True, timeout=10)
    except sqlite3.Error as e:
        print(f"[ERROR] No se pudo abrir '{ruta}': {e}")
        return

    if args:
        ejecutar(con, " ".join(args))
    else:
        print(f"Conectado a '{ruta}' (solo lectura). Escribe 'salir' para terminar.")
        print("Tablas: destinatarios, lotes, paginas\n")
        while True:
            sql = input("sql> ").strip().rstrip(";")
            if sql.lower() in ("salir", "exit", "quit", ".quit"):
                break
            if sql:
                ejecutar(con, sql)
    con.close()


if __name__ == "__main__":
    main()
