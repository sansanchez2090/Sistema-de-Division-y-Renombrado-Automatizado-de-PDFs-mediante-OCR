"""
main.py - Punto de arranque.
Por ahora es un menú de consola para probar la lógica con prints.
En la Fase 2 este archivo solo abrirá la ventana de src/interfaz.py.

Ejecutar desde la carpeta DivisionPDFs:  python main.py
"""

from src import procesador as p

EXCEL_POR_DEFECTO = "BasePruebaDivision.xlsx"
PDF_POR_DEFECTO = "InvitacionesCongreMinDir.pdf"
CARPETA_LOTES_POR_DEFECTO = "lotes"


# ----------------------------------------------------------------------
# Utilidades de entrada
# ----------------------------------------------------------------------
def pedir_ruta(mensaje, por_defecto):
    texto = input(f"{mensaje} [{por_defecto}]: ").strip().strip('"')
    return texto or por_defecto


def pedir_excels():
    entrada = input(
        f"Ruta(s) de Excel separadas por coma [{EXCEL_POR_DEFECTO}]: "
    ).strip()
    return [r.strip().strip('"') for r in entrada.split(",") if r.strip()] or [
        EXCEL_POR_DEFECTO
    ]


def elegir_hojas(excel):
    hojas = p.listar_hojas(excel)
    print(f"\nPestañas de '{excel}':")
    for i, hoja in enumerate(hojas, 1):
        print(f"  [{i}] {hoja}")
    seleccion = (
        input("¿Cuáles usar? (números separados por comas, o 'todas'): ")
        .strip()
        .lower()
    )
    if seleccion == "todas":
        return hojas
    indices = [int(x) - 1 for x in seleccion.split(",") if x.strip().isdigit()]
    return [hojas[i] for i in indices if 0 <= i < len(hojas)]


def mostrar_estado():
    estado = p.resumen_estado()
    print(
        f"\nEstado actual -> Pendientes: {estado.get('Pendiente', 0)}"
        f" | Procesados: {estado.get('Procesado', 0)}"
    )


# ----------------------------------------------------------------------
# Opciones del menú
# ----------------------------------------------------------------------
def opcion_cargar_excel():
    total_nuevos = 0
    for ruta in pedir_excels():
        try:
            hojas = elegir_hojas(ruta)
            res = p.cargar_excel(ruta, hojas)
            total_nuevos += res["nuevos"]
            if res["hojas_con_error"]:
                print(
                    f"[AVISO] Pestañas NO cargadas de '{ruta}': {res['hojas_con_error']}"
                )
        except Exception as e:
            print(f"[ERROR] No se pudo leer '{ruta}': {e}")
    print(f"\n[OK] Destinatarios nuevos cargados: {total_nuevos}")
    mostrar_estado()


def _hay_lista():
    if p.resumen_estado().get("Pendiente") or p.resumen_estado().get("Procesado"):
        return True
    print("[AVISO] La lista maestra está vacía. Carga primero el Excel (opción 1).")
    return False


def opcion_procesar_lote():
    if not _hay_lista():
        return
    pdf = pedir_ruta("Ruta del PDF del lote", PDF_POR_DEFECTO)
    try:
        resumen = p.procesar_lote(pdf)
    except Exception as e:
        print(f"[ERROR] No se pudo procesar el PDF: {e}")
        return
    if resumen:
        p.generar_reporte_control()


def opcion_procesar_carpeta():
    if not _hay_lista():
        return
    carpeta = pedir_ruta("Carpeta con los PDFs", CARPETA_LOTES_POR_DEFECTO)
    try:
        resumenes = p.procesar_carpeta(carpeta)
    except Exception as e:
        print(f"[ERROR] No se pudo leer la carpeta: {e}")
        return
    if resumenes:
        p.generar_reporte_control()


def opcion_reporte_control():
    if not _hay_lista():
        return
    p.generar_reporte_control()


def _mostrar_plan(plan):
    print(f"\n--- Vista previa de la recarga ({plan['excel']}) ---")
    print(f"  Nuevos a agregar:        {len(plan['nuevos'])}")
    print(f"  Registros a actualizar:  {len(plan['actualizar'])}")
    print(f"  Sin cambios:             {plan['sin_cambio']}")
    print(f"  En BD pero no en Excel:  {len(plan['ausentes'])}  (NO se borran)")
    print(f"  Conflictos (omitidos):   {len(plan['conflictos'])}")

    for r in plan["nuevos"][:10]:
        print(f"    + {r['nombre']} (ID {r['id']}, '{r['pestana']}')")
    for r in plan["actualizar"][:10]:
        print(f"    ~ {r['nombre_actual']} ('{r['pestana']}'): {r['cambios']}")
    for texto in plan["ausentes"][:10]:
        print(f"    ? {texto}")
    for texto in plan["conflictos"][:10]:
        print(f"    ! {texto}")
    for texto in plan["hojas_con_error"]:
        print(f"    [ERROR] {texto}")


def opcion_recarga_emergencia():
    if not _hay_lista():
        return
    print("\nRECARGA DE EMERGENCIA: agrega y actualiza, nunca borra ni toca estados.")
    for ruta in pedir_excels():
        try:
            hojas = elegir_hojas(ruta)
            plan = p.planificar_recarga(ruta, hojas)
        except Exception as e:
            print(f"[ERROR] No se pudo leer '{ruta}': {e}")
            continue

        _mostrar_plan(plan)
        if plan["archivo_desconocido"]:
            print(
                f"\n⚠️ El archivo '{plan['excel']}' NO existe en la base de datos."
                " ¿Fue renombrado? Si confirmas, TODOS sus destinatarios se agregarán"
                " como nuevos y podrían duplicarse."
            )
        if not plan["nuevos"] and not plan["actualizar"]:
            print("No hay nada que aplicar.")
            continue
        if input("Escribe CONFIRMAR para aplicar los cambios: ").strip() == "CONFIRMAR":
            p.aplicar_recarga(plan)
        else:
            print("Cancelado, no se modificó nada.")
    mostrar_estado()


def opcion_reiniciar():
    print("\nFIN DE CAMPAÑA. Esto hará, en orden:")
    print("  1) Generar un reporte de control FINAL en 'reportes/'")
    print("  2) Guardar un respaldo de la base de datos")
    print("  3) VACIAR la base de datos (incluida la lista de destinatarios)")
    print("  4) Borrar los PDFs de 'pdfs_divididos/'")
    if input("Escribe REINICIAR para confirmar: ").strip() != "REINICIAR":
        print("Cancelado.")
        return
    p.reiniciar_campana()
    mostrar_estado()


# ----------------------------------------------------------------------
# Menú principal
# ----------------------------------------------------------------------
def main():
    while True:
        print("\n===== DIVISIÓN DE CARTAS =====")
        mostrar_estado()
        print("  1) Cargar Excel (los ya cargados se ignoran)")
        print("  2) Procesar un lote (un PDF)")
        print("  3) Procesar todos los PDFs de una carpeta")
        print("  4) Generar Excel de control y verificar")
        print("  5) Recargar Excel (emergencia)")
        print("  6) Reiniciar campaña (fin del proceso)")
        print("  7) Salir")
        opcion = input("Opción: ").strip()

        if opcion == "1":
            opcion_cargar_excel()
        elif opcion == "2":
            opcion_procesar_lote()
        elif opcion == "3":
            opcion_procesar_carpeta()
        elif opcion == "4":
            opcion_reporte_control()
        elif opcion == "5":
            opcion_recarga_emergencia()
        elif opcion == "6":
            opcion_reiniciar()
        elif opcion == "7":
            break
        else:
            print("Opción no válida.")


if __name__ == "__main__":
    main()
