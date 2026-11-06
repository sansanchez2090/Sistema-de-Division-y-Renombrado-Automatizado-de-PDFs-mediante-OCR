"""
interfaz.py - Ventana principal (CustomTkinter).

Cómo se comunican la ventana y el trabajo pesado:
  - La ventana NUNCA ejecuta OCR ni toca la base de datos en procesos largos.
    Cada tarea corre en un hilo aparte (threading.Thread).
  - El hilo NO toca widgets. Solo pone mensajes en una cola (queue.Queue):
        ("log", texto)                -> línea para el registro de actividad
        ("progreso", actual, total)   -> avance de la barra
        ("fin", resultado, error)     -> la tarea terminó
  - La ventana revisa la cola cada 100 ms con after() y actualiza los widgets.
  - "Cancelar" activa un threading.Event que procesar_lote() revisa entre páginas.
"""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from . import procesador as p

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

INTERVALO_COLA_MS = 100
MAX_LINEAS_LOG = 3000
COLOR_PELIGRO = "#B3261E"
COLOR_PELIGRO_HOVER = "#8C1D18"
COLOR_OK = "#1B7F3B"
COLOR_AVISO = "#B26A00"
COLOR_NEUTRO = "#444444"


# ----------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------
def _limpiar(texto):
    """Quita emojis fuera del rango básico: Tk puede fallar o dibujarlos como cuadros."""
    return "".join(ch for ch in str(texto) if ord(ch) <= 0xFFFF)


def _abrir_carpeta(ruta):
    os.makedirs(ruta, exist_ok=True)
    ruta = os.path.abspath(ruta)
    if hasattr(os, "startfile"):  # Windows
        os.startfile(ruta)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", ruta])
    else:
        subprocess.Popen(["xdg-open", ruta])


def _texto_verificacion(v):
    """Convierte el dict de verificar_completitud() en (texto, tipo) para un mensaje."""
    lineas = [
        f"Destinatarios en la lista: {v['total_lista']}",
        f"Páginas procesadas: {v['paginas_procesadas']}",
        f"Estado del proceso: {v['estado']}",
    ]
    terminado = v["estado"] != "EN CURSO"
    alerta = False

    if v["estado"] == "EN CURSO":
        lineas.append(f"Faltan {-v['diferencia']} carta(s) por procesar.")
    elif v["estado"] == "INCONSISTENTE":
        lineas.append(
            f"Hay {v['diferencia']} página(s) de más. ¿Se repitió un PDF o hay cartas de más?"
        )
        alerta = True

    if v["paginas_por_revisar"]:
        lineas.append(
            f"\n⚠ {v['paginas_por_revisar']} página(s) requieren revisión manual"
            " (pestaña 'Por revisar' del Excel de control)."
        )
        alerta = alerta or terminado
    if v["destinatarios_pendientes"]:
        lineas.append(
            f"\n⚠ {v['destinatarios_pendientes']} destinatario(s) aún sin carta asignada."
        )
        alerta = alerta or terminado
    if terminado and not alerta:
        lineas.append(
            "\n✔ Todo en orden: cada destinatario tiene su carta y no hay nada por revisar."
        )

    return "\n".join(lineas), ("warning" if alerta else "info")


# ----------------------------------------------------------------------
# Aplicación
# ----------------------------------------------------------------------
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("División y renombrado de cartas")
        # El tamaño se adapta a la pantalla: 1366x768 y 1920x1080 funcionan igual de bien.
        # CustomTkinter escala las medidas según el zoom de Windows, por eso se divide.
        try:
            escala = self._get_window_scaling()
        except Exception:
            escala = 1.0
        ancho = min(1020, int(self.winfo_screenwidth() / escala) - 40)
        alto = min(
            820, int(self.winfo_screenheight() / escala) - 110
        )  # barra de tareas y título
        self.geometry(f"{ancho}x{alto}")
        self.minsize(min(860, ancho), min(540, alto))

        self.cola = queue.Queue()
        self.detener = threading.Event()
        self.hilo = None
        self._cerrando = False
        self._after_id = None
        self.excels = []  # un dict por Excel agregado: {"ruta", "bloque", "vars"}
        self.botones = []  # botones que se bloquean mientras corre una tarea

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=4)  # pestañas: es lo que más espacio necesita
        self.grid_rowconfigure(4, weight=1, minsize=90)  # registro de actividad

        self._construir_encabezado()
        self._construir_estado()
        self._construir_tabs()
        self._construir_progreso()
        self._construir_registro()

        self.protocol("WM_DELETE_WINDOW", self._al_cerrar)
        self._actualizar_estado()
        self._after_id = self.after(INTERVALO_COLA_MS, self._procesar_cola)

    # ------------------------------------------------------------------
    # Construcción de la ventana
    # ------------------------------------------------------------------
    def _construir_encabezado(self):
        ctk.CTkLabel(
            self,
            text="División y renombrado de cartas",
            font=ctk.CTkFont(size=22, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=18, pady=(10, 2))

    def _construir_estado(self):
        marco = ctk.CTkFrame(self)
        marco.grid(row=1, column=0, sticky="ew", padx=16, pady=(4, 6))
        self.valores = {}
        tarjetas = [
            ("total", "Destinatarios en lista"),
            ("con_carta", "Con carta asignada"),
            ("pendientes", "Pendientes"),
            ("por_revisar", "Páginas por revisar"),
            ("estado", "Estado del proceso"),
        ]
        for i, (clave, titulo) in enumerate(tarjetas):
            marco.grid_columnconfigure(i, weight=1)
            caja = ctk.CTkFrame(marco, fg_color="transparent")
            caja.grid(row=0, column=i, padx=6, pady=4)
            tam = 16 if clave == "estado" else 20
            valor = ctk.CTkLabel(
                caja, text="-", font=ctk.CTkFont(size=tam, weight="bold")
            )
            valor.pack()
            ctk.CTkLabel(
                caja, text=titulo, font=ctk.CTkFont(size=11), text_color=COLOR_NEUTRO
            ).pack()
            self.valores[clave] = valor

    def _construir_tabs(self):
        self.tabs = ctk.CTkTabview(self)
        self.tabs.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 8))
        self._tab_cargar(self.tabs.add("1. Cargar Excel"))
        self._tab_procesar(self.tabs.add("2. Procesar cartas"))
        self._tab_reportes(self.tabs.add("3. Reportes y cierre"))

    def _construir_progreso(self):
        marco = ctk.CTkFrame(self, fg_color="transparent")
        marco.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 4))
        marco.grid_columnconfigure(0, weight=1)
        self.lbl_progreso = ctk.CTkLabel(marco, text="Listo.", anchor="w")
        self.lbl_progreso.grid(row=0, column=0, sticky="w")
        self.barra = ctk.CTkProgressBar(marco)
        self.barra.grid(row=1, column=0, sticky="ew", pady=(2, 0))
        self.barra.set(0)
        self.btn_cancelar = ctk.CTkButton(
            marco,
            text="Cancelar",
            width=110,
            fg_color=COLOR_PELIGRO,
            hover_color=COLOR_PELIGRO_HOVER,
            state="disabled",
            command=self._cancelar,
        )
        self.btn_cancelar.grid(row=0, column=1, rowspan=2, padx=(12, 0))

    def _construir_registro(self):
        self.txt_log = ctk.CTkTextbox(
            self, height=110, wrap="word", font=ctk.CTkFont(family="Consolas", size=12)
        )
        self.txt_log.grid(row=4, column=0, sticky="nsew", padx=16, pady=(0, 4))
        self.txt_log.configure(state="disabled")
        ctk.CTkLabel(
            self,
            text=f"Carpeta de trabajo: {os.getcwd()}",
            font=ctk.CTkFont(size=11),
            text_color=COLOR_NEUTRO,
            anchor="w",
        ).grid(row=5, column=0, sticky="ew", padx=18, pady=(0, 8))

    def _contenedor_desplazable(self, tab):
        """Si el contenido no cabe, aparece una barra de desplazamiento en vez de cortarse."""
        marco = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        marco.pack(fill="both", expand=True)
        return marco

    # --- Pestaña 1 ----------------------------------------------------
    def _tab_cargar(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(2, weight=1)  # solo la lista crece o se encoge

        ctk.CTkLabel(
            tab,
            text=(
                "Agrega los Excel con los destinatarios y marca las pestañas que se van a usar.\n"
                "Hazlo al iniciar cada campaña, para trabajar con la última versión."
            ),
            justify="left",
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=8, pady=(2, 6))

        self.btn_agregar = ctk.CTkButton(
            tab, text="Agregar Excel...", width=170, command=self._agregar_excel
        )
        self.btn_agregar.grid(row=1, column=0, sticky="w", padx=8)

        self.frame_excels = ctk.CTkScrollableFrame(tab, height=60)
        self.frame_excels.grid(row=2, column=0, sticky="nsew", padx=8, pady=6)

        self.btn_cargar = ctk.CTkButton(
            tab,
            text="Cargar lista de destinatarios",
            height=40,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._cargar_lista,
        )
        self.btn_cargar.grid(row=3, column=0, sticky="ew", padx=8, pady=(0, 4))
        self.botones += [self.btn_agregar, self.btn_cargar]

    def _agregar_excel(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el Excel de destinatarios",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Todos los archivos", "*.*")],
        )
        if not ruta:
            return
        nombre = os.path.basename(ruta)
        if any(
            os.path.basename(e["ruta"]).lower() == nombre.lower() for e in self.excels
        ):
            messagebox.showwarning(
                "Archivo repetido",
                f"Ya agregaste un archivo llamado '{nombre}'.\n\n"
                "El programa identifica cada Excel por su nombre, así que no puede haber dos iguales.",
            )
            return
        try:
            hojas = p.listar_hojas(ruta)
        except Exception as e:
            messagebox.showerror("No se pudo leer el Excel", f"{e}")
            return

        bloque = ctk.CTkFrame(self.frame_excels)
        bloque.pack(fill="x", padx=4, pady=6)
        entrada = {"ruta": ruta, "bloque": bloque, "vars": {}}

        cabecera = ctk.CTkFrame(bloque, fg_color="transparent")
        cabecera.pack(fill="x", padx=8, pady=(8, 2))
        ctk.CTkLabel(
            cabecera, text=nombre, font=ctk.CTkFont(size=14, weight="bold")
        ).pack(side="left")
        ctk.CTkButton(
            cabecera,
            text="Quitar",
            width=70,
            fg_color="gray50",
            hover_color="gray35",
            command=lambda e=entrada: self._quitar_excel(e),
        ).pack(side="right")

        cuadricula = ctk.CTkFrame(bloque, fg_color="transparent")
        cuadricula.pack(fill="x", padx=8, pady=(0, 8))
        for i, hoja in enumerate(hojas):
            var = tk.BooleanVar(value=True)
            entrada["vars"][hoja] = var
            ctk.CTkCheckBox(cuadricula, text=hoja, variable=var).grid(
                row=i // 3, column=i % 3, sticky="w", padx=8, pady=3
            )
        self.excels.append(entrada)

    def _quitar_excel(self, entrada):
        entrada["bloque"].destroy()
        if entrada in self.excels:
            self.excels.remove(entrada)

    def _cargar_lista(self):
        seleccion = []
        for e in self.excels:
            hojas = [h for h, v in e["vars"].items() if v.get()]
            if hojas:
                seleccion.append((e["ruta"], hojas))
        if not seleccion:
            messagebox.showinfo(
                "Falta información",
                "Agrega al menos un Excel y marca al menos una pestaña.",
            )
            return
        self._ejecutar("Cargando lista de destinatarios", self._tarea_cargar, seleccion)

    def _tarea_cargar(self, log, seleccion):
        nuevos = problemas = 0
        for ruta, hojas in seleccion:
            log(f"\nArchivo: {os.path.basename(ruta)}")
            r = p.cargar_excel(ruta, hojas, log=log)
            nuevos += r["nuevos"]
            problemas += len(r["hojas_con_error"]) + r["conflictos"]
        mensaje = f"Se cargaron {nuevos} destinatarios nuevos."
        if problemas:
            mensaje += (
                f"\n\n⚠ Hubo {problemas} problema(s): pestañas que no se cargaron o nombres en conflicto."
                " Revisa el registro de actividad."
            )
        return ("Carga terminada", mensaje, "warning" if problemas else "info")

    # --- Pestaña 2 ----------------------------------------------------
    def _tab_procesar(self, tab):
        tab = self._contenedor_desplazable(tab)
        ctk.CTkLabel(
            tab,
            text=(
                "Selecciona el PDF escaneado (o una carpeta con varios PDFs). Cada PDF es un lote.\n"
                "Si un lote se interrumpe, vuelve a seleccionar el mismo PDF y continuará donde quedó."
            ),
            justify="left",
            anchor="w",
        ).pack(fill="x", padx=8, pady=(2, 8))

        self.seg_modo = ctk.CTkSegmentedButton(
            tab, values=["Un PDF", "Carpeta de PDFs"], command=self._cambiar_modo
        )
        self.seg_modo.set("Un PDF")
        self.seg_modo.pack(anchor="w", padx=8, pady=(0, 8))

        fila = ctk.CTkFrame(tab, fg_color="transparent")
        fila.pack(fill="x", padx=8)
        fila.grid_columnconfigure(0, weight=1)
        self.var_ruta = tk.StringVar()
        ctk.CTkEntry(fila, textvariable=self.var_ruta).grid(
            row=0, column=0, sticky="ew"
        )
        self.btn_examinar = ctk.CTkButton(
            fila, text="Examinar...", width=120, command=self._examinar
        )
        self.btn_examinar.grid(row=0, column=1, padx=(8, 0))

        self.btn_iniciar = ctk.CTkButton(
            tab,
            text="Iniciar procesamiento",
            height=40,
            font=ctk.CTkFont(size=16, weight="bold"),
            command=self._iniciar_procesamiento,
        )
        self.btn_iniciar.pack(fill="x", padx=8, pady=(12, 8))

        ctk.CTkLabel(
            tab,
            text="Los PDFs resultantes quedan en 'pdfs_divididos'. Al terminar se actualiza el Excel de control en 'reportes'.",
            justify="left",
            anchor="w",
            text_color=COLOR_NEUTRO,
        ).pack(fill="x", padx=8)
        self.botones += [self.btn_examinar, self.btn_iniciar]

    def _cambiar_modo(self, _valor):
        self.var_ruta.set("")

    def _examinar(self):
        if self.seg_modo.get() == "Un PDF":
            ruta = filedialog.askopenfilename(
                title="Selecciona el PDF escaneado",
                filetypes=[("PDF", "*.pdf"), ("Todos los archivos", "*.*")],
            )
        else:
            ruta = filedialog.askdirectory(title="Selecciona la carpeta con los PDFs")
        if ruta:
            self.var_ruta.set(ruta)

    def _iniciar_procesamiento(self):
        ruta = self.var_ruta.get().strip().strip('"')
        if not ruta or not os.path.exists(ruta):
            messagebox.showwarning(
                "Falta información",
                "Selecciona el PDF o la carpeta que quieres procesar.",
            )
            return
        estado = p.resumen_estado()
        if not (estado.get("Pendiente") or estado.get("Procesado")):
            messagebox.showwarning(
                "Lista vacía",
                "Primero carga la lista de destinatarios en la pestaña '1. Cargar Excel'.",
            )
            return
        self._ejecutar(
            "Procesando cartas",
            self._tarea_procesar,
            ruta,
            self.seg_modo.get(),
            cancelable=True,
        )

    def _tarea_procesar(self, log, ruta, modo):
        if modo == "Un PDF":
            muestra = ruta
        else:
            muestra = next(
                (
                    os.path.join(ruta, f)
                    for f in sorted(os.listdir(ruta))
                    if f.lower().endswith(".pdf")
                ),
                None,
            )
        problemas = p.verificar_dependencias(muestra)
        if problemas:
            for texto in problemas:
                log(f"[ERROR] {texto}")
            return ("No se puede procesar", "\n\n".join(problemas), "error")

        if modo == "Un PDF":
            r = p.procesar_lote(
                ruta, log=log, progreso=self._progreso, detener=self.detener
            )
            resumenes = [r] if r else []
        else:
            resumenes = p.procesar_carpeta(
                ruta, log=log, progreso=self._progreso, detener=self.detener
            )

        if not resumenes:
            return (
                "Nada que procesar",
                "No se procesó ningún PDF nuevo.\n\nSi ya lo habías procesado, el programa lo "
                "reconoce y no lo repite. Revisa también el registro de actividad.",
                "info",
            )

        reporte = p.generar_reporte_control(log=log)
        texto, tipo = _texto_verificacion(p.verificar_completitud())
        cancelado = any(r["cancelado"] for r in resumenes)

        mensaje = f"Lotes procesados en esta ejecución: {len(resumenes)}\n\n{texto}"
        if cancelado:
            mensaje += (
                "\n\nEl procesamiento se detuvo antes de terminar. Vuelve a seleccionar el mismo PDF"
                " para continuar donde quedó."
            )
            tipo = "warning"
        if reporte is None:
            mensaje += (
                "\n\n⚠ No se pudo actualizar el Excel de control (¿está abierto?). Ciérralo y usa"
                " 'Generar Excel de control' en la pestaña 3."
            )
            tipo = "warning"
        return (
            "Proceso detenido" if cancelado else "Procesamiento terminado",
            mensaje,
            tipo,
        )

    # --- Pestaña 3 ----------------------------------------------------
    def _tab_reportes(self, tab):
        tab = self._contenedor_desplazable(tab)
        ctk.CTkLabel(
            tab,
            text="Verificación y reportes",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w",
        ).pack(fill="x", padx=8, pady=(4, 4))
        ctk.CTkLabel(
            tab,
            text=(
                "El Excel de control tiene 4 pestañas: Resumen, Destinatarios, Páginas y Por revisar.\n"
                "Ciérralo en Excel antes de volver a generarlo."
            ),
            justify="left",
            anchor="w",
            text_color=COLOR_NEUTRO,
        ).pack(fill="x", padx=8, pady=(0, 8))

        fila = ctk.CTkFrame(tab, fg_color="transparent")
        fila.pack(fill="x", padx=8)
        self.btn_reporte = ctk.CTkButton(
            fila,
            text="Generar Excel de control y verificar",
            height=38,
            command=self._generar_reporte,
        )
        self.btn_reporte.pack(side="left")
        ctk.CTkButton(
            fila,
            text="Abrir carpeta de reportes",
            height=38,
            fg_color="gray50",
            hover_color="gray35",
            command=lambda: self._abrir(p.CARPETA_REPORTES),
        ).pack(side="left", padx=8)
        ctk.CTkButton(
            fila,
            text="Abrir carpeta de PDFs",
            height=38,
            fg_color="gray50",
            hover_color="gray35",
            command=lambda: self._abrir(p.CARPETA_SALIDA),
        ).pack(side="left")

        ctk.CTkLabel(
            tab,
            text="Opciones avanzadas",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w",
        ).pack(fill="x", padx=8, pady=(16, 4))

        self.btn_recarga = ctk.CTkButton(
            tab,
            text="Recargar un Excel (emergencia)...",
            height=38,
            width=300,
            fg_color="gray50",
            hover_color="gray35",
            command=self._recarga_emergencia,
        )
        self.btn_recarga.pack(anchor="w", padx=8, pady=(0, 4))
        ctk.CTkLabel(
            tab,
            text="Solo si cambió algo en el Excel durante la campaña. Agrega y actualiza; nunca borra.",
            anchor="w",
            text_color=COLOR_NEUTRO,
        ).pack(fill="x", padx=8, pady=(0, 10))

        self.btn_reiniciar = ctk.CTkButton(
            tab,
            text="Finalizar campaña (reiniciar)...",
            height=38,
            width=300,
            fg_color=COLOR_PELIGRO,
            hover_color=COLOR_PELIGRO_HOVER,
            command=self._reiniciar,
        )
        self.btn_reiniciar.pack(anchor="w", padx=8, pady=(0, 4))
        ctk.CTkLabel(
            tab,
            text="Guarda un reporte final y un respaldo, y deja todo vacío para la siguiente campaña.",
            anchor="w",
            text_color=COLOR_NEUTRO,
        ).pack(fill="x", padx=8)
        self.botones += [self.btn_reporte, self.btn_recarga, self.btn_reiniciar]

    def _abrir(self, carpeta):
        try:
            _abrir_carpeta(carpeta)
        except Exception as e:
            messagebox.showerror("No se pudo abrir la carpeta", f"{e}")

    def _generar_reporte(self):
        if p.verificar_completitud()["total_lista"] == 0:
            messagebox.showinfo(
                "Sin datos", "La lista está vacía. Carga primero el Excel."
            )
            return
        self._ejecutar("Generando Excel de control", self._tarea_reporte)

    def _tarea_reporte(self, log):
        ruta = p.generar_reporte_control(log=log)
        if ruta is None:
            return (
                "No se pudo generar el reporte",
                "El archivo del reporte parece estar abierto en Excel.\nCiérralo e inténtalo de nuevo.",
                "error",
            )
        texto, tipo = _texto_verificacion(p.verificar_completitud())
        return (
            "Reporte generado",
            f"Se guardó en:\n{os.path.abspath(ruta)}\n\n{texto}",
            tipo,
        )

    # --- Recarga de emergencia ---------------------------------------
    def _recarga_emergencia(self):
        if p.verificar_completitud()["total_lista"] == 0:
            messagebox.showinfo(
                "Sin datos",
                "La lista está vacía. Usa la pestaña 1 para cargar el Excel.",
            )
            return
        ruta = filedialog.askopenfilename(
            title="Selecciona el Excel actualizado",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Todos los archivos", "*.*")],
        )
        if not ruta:
            return
        try:
            hojas = p.listar_hojas(ruta)
        except Exception as e:
            messagebox.showerror("No se pudo leer el Excel", f"{e}")
            return
        elegidas = self._dialogo_hojas(ruta, hojas)
        if not elegidas:
            return
        self._ejecutar(
            "Comparando el Excel con la base de datos",
            self._tarea_planificar,
            ruta,
            elegidas,
        )

    def _dialogo_hojas(self, ruta, hojas):
        top = ctk.CTkToplevel(self)
        top.title("Pestañas a revisar")
        top.geometry("420x400")
        top.transient(self)
        ctk.CTkLabel(
            top,
            text=f"Pestañas de {os.path.basename(ruta)}:",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(padx=16, pady=(16, 6), anchor="w")

        marco = ctk.CTkScrollableFrame(top)
        marco.pack(fill="both", expand=True, padx=16, pady=6)
        variables = {}
        for hoja in hojas:
            variables[hoja] = tk.BooleanVar(value=True)
            ctk.CTkCheckBox(marco, text=hoja, variable=variables[hoja]).pack(
                anchor="w", pady=3
            )

        resultado = {"hojas": None}

        def aceptar():
            resultado["hojas"] = [h for h, v in variables.items() if v.get()]
            top.destroy()

        fila = ctk.CTkFrame(top, fg_color="transparent")
        fila.pack(fill="x", padx=16, pady=(6, 16))
        ctk.CTkButton(fila, text="Continuar", command=aceptar).pack(side="right")
        ctk.CTkButton(
            fila,
            text="Cancelar",
            fg_color="gray50",
            hover_color="gray35",
            command=top.destroy,
        ).pack(side="right", padx=8)

        top.after(150, lambda: (top.grab_set(), top.focus_force()))
        self.wait_window(top)
        return resultado["hojas"]

    def _tarea_planificar(self, log, ruta, hojas):
        plan = p.planificar_recarga(ruta, hojas)
        return lambda: self._mostrar_plan(plan)  # se ejecuta en el hilo de la ventana

    def _mostrar_plan(self, plan):
        lineas = [
            f"Archivo: {plan['excel']}",
            "",
            f"Nuevos a agregar: {len(plan['nuevos'])}",
            f"Registros a actualizar: {len(plan['actualizar'])}",
            f"Sin cambios: {plan['sin_cambio']}",
            f"En la base de datos pero no en el Excel: {len(plan['ausentes'])} (NO se borran)",
            f"Conflictos (se omiten): {len(plan['conflictos'])}",
        ]
        if plan["archivo_desconocido"]:
            lineas += [
                "",
                "⚠ Este archivo NO existe en la base de datos. ¿Fue renombrado? Si aplicas los cambios,"
                " TODOS sus destinatarios se agregarán como nuevos y podrían duplicarse.",
            ]

        def seccion(titulo, items):
            if not items:
                return
            lineas.append("")
            lineas.append(titulo)
            lineas.extend(f"  {x}" for x in items[:15])
            if len(items) > 15:
                lineas.append(f"  ... y {len(items) - 15} más")

        seccion(
            "Nuevos:",
            [f"+ {r['nombre']} (ID {r['id']}, {r['pestana']})" for r in plan["nuevos"]],
        )
        seccion(
            "Cambios:",
            [
                f"~ {r['nombre_actual']} ({r['pestana']}): "
                + ", ".join(f"{k} -> {v or '(vacío)'}" for k, v in r["cambios"].items())
                for r in plan["actualizar"]
            ],
        )
        seccion(
            "En la base de datos pero no en el Excel:",
            [f"? {x}" for x in plan["ausentes"]],
        )
        seccion("Conflictos:", [f"! {x}" for x in plan["conflictos"]])
        seccion("Errores en el Excel:", plan["hojas_con_error"])

        hay_cambios = bool(plan["nuevos"] or plan["actualizar"])
        top = ctk.CTkToplevel(self)
        top.title("Vista previa de la recarga")
        top.geometry("720x560")
        top.transient(self)
        ctk.CTkLabel(
            top,
            text=(
                "Revisa los cambios antes de aplicarlos"
                if hay_cambios
                else "No hay nada que aplicar"
            ),
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(padx=16, pady=(16, 6), anchor="w")
        caja = ctk.CTkTextbox(top, wrap="word")
        caja.pack(fill="both", expand=True, padx=16, pady=6)
        caja.insert("1.0", _limpiar("\n".join(lineas)))
        caja.configure(state="disabled")

        def aplicar():
            if plan["archivo_desconocido"] and not messagebox.askyesno(
                "Confirmar",
                "Este archivo no estaba en la base de datos. ¿Seguro que quieres aplicar?",
                parent=top,
            ):
                return
            top.destroy()
            self._ejecutar("Aplicando la recarga", self._tarea_aplicar, plan)

        fila = ctk.CTkFrame(top, fg_color="transparent")
        fila.pack(fill="x", padx=16, pady=(6, 16))
        ctk.CTkButton(
            fila,
            text="Aplicar cambios",
            state="normal" if hay_cambios else "disabled",
            command=aplicar,
        ).pack(side="right")
        ctk.CTkButton(
            fila,
            text="Cerrar",
            fg_color="gray50",
            hover_color="gray35",
            command=top.destroy,
        ).pack(side="right", padx=8)
        top.after(150, lambda: (top.grab_set(), top.focus_force()))

    def _tarea_aplicar(self, log, plan):
        r = p.aplicar_recarga(plan, log=log)
        return (
            "Recarga aplicada",
            f"Se agregaron {r['insertados']} destinatarios y se actualizaron {r['actualizados']}.",
            "info",
        )

    # --- Reinicio -----------------------------------------------------
    def _reiniciar(self):
        v = p.verificar_completitud()
        if v["total_lista"] == 0:
            messagebox.showinfo("Nada que reiniciar", "La base de datos ya está vacía.")
            return
        mensaje = (
            "Al finalizar la campaña, el programa hará, en orden:\n\n"
            "  1. Generar un reporte de control FINAL en la carpeta 'reportes'.\n"
            "  2. Guardar un respaldo de la base de datos.\n"
            "  3. VACIAR la base de datos, incluida la lista de destinatarios.\n"
            "  4. Borrar los PDFs de la carpeta 'pdfs_divididos'.\n\n"
            f"Estado actual: {v['estado']}. Destinatarios sin carta: {v['destinatarios_pendientes']}. "
            f"Páginas por revisar: {v['paginas_por_revisar']}.\n\n"
            "Tus PDFs originales no se tocan."
        )
        if self._confirmar_con_palabra("Finalizar campaña", mensaje, "REINICIAR"):
            self._ejecutar("Finalizando campaña", self._tarea_reiniciar)

    def _confirmar_con_palabra(self, titulo, mensaje, palabra):
        top = ctk.CTkToplevel(self)
        top.title(titulo)
        top.geometry("540x430")
        top.transient(self)
        ctk.CTkLabel(
            top, text=mensaje, justify="left", wraplength=490, anchor="w"
        ).pack(padx=20, pady=(20, 10), anchor="w")
        ctk.CTkLabel(top, text=f"Escribe {palabra} para confirmar:", anchor="w").pack(
            padx=20, anchor="w"
        )
        var = tk.StringVar()
        ctk.CTkEntry(top, textvariable=var).pack(fill="x", padx=20, pady=8)

        resultado = {"ok": False}

        def confirmar():
            resultado["ok"] = True
            top.destroy()

        fila = ctk.CTkFrame(top, fg_color="transparent")
        fila.pack(fill="x", padx=20, pady=(6, 16))
        boton = ctk.CTkButton(
            fila,
            text="Confirmar",
            state="disabled",
            fg_color=COLOR_PELIGRO,
            hover_color=COLOR_PELIGRO_HOVER,
            command=confirmar,
        )
        boton.pack(side="right")
        ctk.CTkButton(
            fila,
            text="Cancelar",
            fg_color="gray50",
            hover_color="gray35",
            command=top.destroy,
        ).pack(side="right", padx=8)
        var.trace_add(
            "write",
            lambda *_: boton.configure(
                state="normal" if var.get().strip() == palabra else "disabled"
            ),
        )

        top.after(150, lambda: (top.grab_set(), top.focus_force()))
        self.wait_window(top)
        return resultado["ok"]

    def _tarea_reiniciar(self, log):
        ok = p.reiniciar_campana(log=log)
        if ok:
            return (
                "Campaña finalizada",
                "Listo. El reporte final y el respaldo quedaron guardados.\n"
                "La base de datos está vacía para la siguiente campaña.",
                "info",
            )
        return (
            "No se reinició",
            "No se pudo completar el reinicio y no se borró nada. Revisa el registro de actividad.",
            "error",
        )

    # ------------------------------------------------------------------
    # Hilo de trabajo y cola de mensajes
    # ------------------------------------------------------------------
    def _ejecutar(self, titulo, funcion, *args, cancelable=False):
        """Ejecuta funcion(log, *args) en un hilo. La ventana sigue respondiendo."""
        if self.hilo is not None and self.hilo.is_alive():
            return
        self.detener.clear()
        self._bloquear(True, cancelable)
        self.barra.set(0)
        self.lbl_progreso.configure(text=titulo + "...")
        self._log(f"\n=== {titulo} ===")

        def trabajo():
            try:
                resultado = funcion(lambda m: self.cola.put(("log", m)), *args)
                self.cola.put(("fin", resultado, None))
            except Exception as e:  # cualquier error llega a la ventana, no se pierde
                self.cola.put(("fin", None, e))

        self.hilo = threading.Thread(target=trabajo, daemon=True)
        self.hilo.start()

    def _progreso(self, actual, total):
        """Callback que usa el procesador (corre en el hilo de trabajo)."""
        self.cola.put(("progreso", actual, total))

    def _procesar_cola(self):
        # Se reprograma primero: si un cuadro de diálogo detiene esta función, el ciclo sigue vivo.
        self._after_id = self.after(INTERVALO_COLA_MS, self._procesar_cola)
        try:
            while True:
                msg = self.cola.get_nowait()
                if msg[0] == "log":
                    self._log(msg[1])
                elif msg[0] == "progreso":
                    _, actual, total = msg
                    self.barra.set(actual / total if total else 0)
                    self.lbl_progreso.configure(text=f"Página {actual} de {total}")
                elif msg[0] == "fin":
                    self._terminar(msg[1], msg[2])
        except queue.Empty:
            pass

    def _terminar(self, resultado, error):
        self._bloquear(False)
        self._actualizar_estado()
        if self._cerrando:
            self._salir()
            return
        self.lbl_progreso.configure(text="Listo.")

        if error is not None:
            self._log(f"[ERROR] {error}")
            messagebox.showerror(
                "Ocurrió un problema",
                f"{error}\n\nSi el problema continúa, copia el registro de actividad y compártelo con soporte.",
            )
            return
        if callable(resultado):
            resultado()
            return
        if resultado:
            titulo, mensaje, tipo = resultado
            {
                "info": messagebox.showinfo,
                "warning": messagebox.showwarning,
                "error": messagebox.showerror,
            }[tipo](titulo, mensaje)

    def _bloquear(self, ocupado, cancelable=False):
        for boton in self.botones:
            boton.configure(state="disabled" if ocupado else "normal")
        self.btn_cancelar.configure(
            state="normal" if (ocupado and cancelable) else "disabled"
        )

    def _cancelar(self):
        self.detener.set()
        self.btn_cancelar.configure(state="disabled")
        self._log("Cancelando... se detendrá al terminar la página actual.")

    # ------------------------------------------------------------------
    # Registro, estado y cierre
    # ------------------------------------------------------------------
    def _log(self, texto):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", _limpiar(texto) + "\n")
        if int(self.txt_log.index("end-1c").split(".")[0]) > MAX_LINEAS_LOG:
            self.txt_log.delete("1.0", "500.0")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _actualizar_estado(self):
        try:
            v = p.verificar_completitud()
        except Exception as e:
            self._log(f"[AVISO] No se pudo leer el estado: {e}")
            return
        total, pend = v["total_lista"], v["destinatarios_pendientes"]
        self.valores["total"].configure(text=str(total))
        self.valores["con_carta"].configure(text=str(total - pend))
        self.valores["pendientes"].configure(
            text=str(pend), text_color=COLOR_AVISO if pend else COLOR_OK
        )
        self.valores["por_revisar"].configure(
            text=str(v["paginas_por_revisar"]),
            text_color=COLOR_AVISO if v["paginas_por_revisar"] else COLOR_OK,
        )
        if total == 0:
            texto, color = "Sin lista", COLOR_NEUTRO
        else:
            texto = v["estado"]
            color = {
                "EN CURSO": COLOR_AVISO,
                "COMPLETO": COLOR_OK,
                "INCONSISTENTE": COLOR_PELIGRO,
            }[texto]
        self.valores["estado"].configure(text=texto, text_color=color)

    def _al_cerrar(self):
        if self.hilo is not None and self.hilo.is_alive():
            if not messagebox.askyesno(
                "Proceso en curso",
                "Hay un proceso en ejecución. ¿Quieres detenerlo y salir?\n\n"
                "Lo ya procesado queda guardado y podrás continuar después.",
            ):
                return
            self._cerrando = True
            self.detener.set()
            self._log("Cerrando: esperando a que termine la página actual...")
        else:
            self._salir()

    def _salir(self):
        if self._after_id is not None:
            self.after_cancel(self._after_id)
            self._after_id = None
        self.destroy()


def iniciar_app():
    App().mainloop()
