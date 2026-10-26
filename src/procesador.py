"""
procesador.py
Lógica del negocio: lectura de Excel, base de datos SQLite, OCR, cruce de
nombres, renombrado de PDFs, verificación, reporte de control y reinicio.

IMPORTANTE para la futura interfaz:
  - Ninguna función hace input(). Todas reciben parámetros.
  - Todas reciben `log` (por defecto print). Con la interfaz solo cambiamos
    `log` por una función que ponga el mensaje en una cola.
  - procesar_lote()/procesar_carpeta() reciben `progreso` (callback) y
    `detener` (threading.Event).
  - Cada función abre su propia conexión SQLite: es seguro usarlas desde un
    hilo distinto al de la ventana.
"""

import hashlib
import os
import re
import sqlite3
import time
import unicodedata
from datetime import datetime

import pandas as pd
import pytesseract
from openpyxl.styles import Font
from pdf2image import convert_from_path
from pypdf import PdfReader, PdfWriter
from rapidfuzz import fuzz

# ----------------------------------------------------------------------
# Configuración
# ----------------------------------------------------------------------
RUTA_DB = "proyecto.db"
CARPETA_SALIDA = "pdfs_divididos"
CARPETA_REPORTES = "reportes"

COLUMNAS = ("id", "nombre", "cargo", "entidad", "correo")  # Columnas del Excel
CAMPOS_EDITABLES = ("nombre", "cargo", "entidad", "correo")
MAX_FILAS_ENCABEZADO = 30  # Filas donde se busca el encabezado "Nombre"

DPI_OCR = 300
FRACCION_SUPERIOR = 0.45  # OCR solo al 45% superior de la página
UMBRAL_ZONA = 90  # Puntaje mínimo buscando en la zona del destinatario
UMBRAL_PAGINA = 96  # Puntaje mínimo buscando en todo el texto leído (más estricto)
DIF_AMBIGUO = 3  # Si los 2 mejores difieren menos de esto -> REVISAR
MIN_CARACTERES_OCR = 20  # Si el OCR lee menos que esto, la página se considera vacía
FACTOR_LONGITUD = (
    0.8  # La zona debe medir al menos 80% del nombre para poder compararlo
)
# Palabras que anteceden al nombre del destinatario (ya normalizadas, sin tildes).
# Si en el futuro aparece otro tratamiento ("Ingeniero", "General"...), agrégalo aquí.
TRATAMIENTOS = (
    r"(?:senor(?:es|a|as)?|doctor(?:a|es|as)?|dr|dra|honorable|magistrado|magistrada)"
)

# Si Tesseract o Poppler no están en el PATH (Windows), descomenta y ajusta:
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# RUTA_POPPLER = r"C:\poppler\Library\bin"
RUTA_POPPLER = None


# ----------------------------------------------------------------------
# Utilidades de texto y archivos
# ----------------------------------------------------------------------
def normalizar_texto(texto):
    """Minúsculas, sin tildes, sin puntuación y con espacios colapsados."""
    if not isinstance(texto, str):
        return ""
    texto = unicodedata.normalize("NFD", texto.lower())
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = re.sub(r"[^a-z0-9\s]", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def nombre_archivo_seguro(nombre):
    return "".join(
        c for c in nombre if c.isalnum() or c in (" ", "_", "-", ".")
    ).strip()


def ruta_disponible(carpeta, base, extension=".pdf"):
    """Devuelve una ruta que no exista, agregando _2, _3... si hace falta."""
    ruta = os.path.join(carpeta, f"{base}{extension}")
    contador = 2
    while os.path.exists(ruta):
        ruta = os.path.join(carpeta, f"{base}_{contador}{extension}")
        contador += 1
    return ruta


def calcular_hash(ruta, bloque=1024 * 1024):
    """Huella SHA-256 del archivo: identifica un PDF aunque cambie de nombre."""
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for trozo in iter(lambda: f.read(bloque), b""):
            h.update(trozo)
    return h.hexdigest()


def _valor_celda(fila, posicion):
    if posicion is None:
        return ""
    v = fila.iloc[posicion]
    return v.strip() if isinstance(v, str) else ""


def _normalizar_id(valor):
    """'12', ' 12 ' y '12.0' se tratan como el mismo ID."""
    valor = (valor or "").strip()
    return re.sub(r"\.0+$", "", valor)


# ----------------------------------------------------------------------
# Base de datos (SQLite)
# ----------------------------------------------------------------------
ESQUEMA = """
CREATE TABLE IF NOT EXISTS destinatarios (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    excel_origen  TEXT NOT NULL,
    pestana       TEXT NOT NULL,
    id_excel      TEXT NOT NULL,
    nombre        TEXT NOT NULL,
    nombre_norm   TEXT NOT NULL UNIQUE,
    cargo         TEXT,
    entidad       TEXT,
    correo        TEXT,
    estado        TEXT NOT NULL DEFAULT 'Pendiente',
    archivo       TEXT,
    lote_id       INTEGER,
    UNIQUE (excel_origen, pestana, id_excel)
);

CREATE TABLE IF NOT EXISTS lotes (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha          TEXT,
    pdf_origen     TEXT,
    hash_pdf       TEXT,
    total_paginas  INTEGER,
    estado         TEXT NOT NULL DEFAULT 'EN_PROCESO',
    duracion_seg   REAL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS paginas (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    lote_id          INTEGER NOT NULL REFERENCES lotes(id),
    num_pagina       INTEGER NOT NULL,
    destinatario_id  INTEGER REFERENCES destinatarios(id),
    puntaje          REAL,
    estado           TEXT NOT NULL,
    archivo          TEXT,
    texto_ocr        TEXT
);
"""


def conectar(ruta_db=RUTA_DB):
    """Abre (o crea) la base de datos y garantiza que las tablas existan."""
    con = sqlite3.connect(ruta_db)
    con.row_factory = sqlite3.Row  # permite fila["nombre"] en vez de fila[0]
    columnas = [r["name"] for r in con.execute("PRAGMA table_info(destinatarios)")]
    if columnas and "id_excel" not in columnas:
        con.close()
        raise RuntimeError(
            f"'{ruta_db}' es de una versión anterior del programa. "
            "Muévela o bórrala y vuelve a cargar el Excel."
        )
    con.executescript(ESQUEMA)
    return con


def resumen_estado(ruta_db=RUTA_DB):
    """Devuelve {'Pendiente': n, 'Procesado': m} con el estado actual."""
    con = conectar(ruta_db)
    try:
        filas = con.execute(
            "SELECT estado, COUNT(*) AS n FROM destinatarios GROUP BY estado"
        ).fetchall()
        return {f["estado"]: f["n"] for f in filas}
    finally:
        con.close()


# ----------------------------------------------------------------------
# Lectura del Excel
# ----------------------------------------------------------------------
def listar_hojas(excel_path):
    return pd.ExcelFile(excel_path).sheet_names


def leer_destinatarios_hoja(excel_path, hoja):
    """
    Busca la fila de encabezado (la que contiene 'Nombre') en las primeras
    filas y devuelve una lista de dicts con id, nombre, cargo, entidad y correo.
    Devuelve None si no encuentra la columna Nombre.
    """
    crudo = pd.read_excel(excel_path, sheet_name=hoja, header=None, dtype=str)

    for fila in range(min(MAX_FILAS_ENCABEZADO, len(crudo))):
        encabezados = [normalizar_texto(v) for v in crudo.iloc[fila]]
        if "nombre" not in encabezados:
            continue

        posiciones = {c: encabezados.index(c) for c in COLUMNAS if c in encabezados}
        registros = []
        for _, datos in crudo.iloc[fila + 1 :].iterrows():
            registro = {c: _valor_celda(datos, posiciones.get(c)) for c in COLUMNAS}
            registro["id"] = _normalizar_id(registro["id"])
            if registro["nombre"]:
                registros.append(registro)
        return registros

    return None


def _validar_registros(registros):
    """Devuelve una lista de errores (vacía si todo está bien)."""
    errores = []
    if all(not r["id"] for r in registros):
        return ["No se encontró la columna 'ID' o está vacía."]

    sin_id = [r["nombre"] for r in registros if not r["id"]]
    if sin_id:
        errores.append(f"{len(sin_id)} destinatario(s) sin ID (ej.: {sin_id[0]}).")

    por_id = {}
    for r in registros:
        if r["id"]:
            por_id.setdefault(r["id"], []).append(r["nombre"])
    for id_, nombres in por_id.items():
        if len(nombres) > 1:
            errores.append(f"ID {id_} repetido: {' / '.join(nombres)}")
    return errores


# ----------------------------------------------------------------------
# Carga inicial de la campaña
# ----------------------------------------------------------------------
def cargar_excel(excel_path, hojas, ruta_db=RUTA_DB, log=print):
    """
    Inserta en la BD los destinatarios de las pestañas indicadas.
    Clave de cada persona: (nombre del archivo Excel, pestaña, ID).
    - Los nuevos quedan como 'Pendiente'.
    - Los que ya existen se ignoran (no se modifica nada).
    - Si una pestaña tiene IDs repetidos o vacíos, NO se carga y se avisa.
    - Si un nombre ya existe bajo otro ID, se avisa y no se inserta.
    """
    excel = os.path.basename(excel_path)
    resultado = {"nuevos": 0, "ya_existian": 0, "conflictos": 0, "hojas_con_error": []}
    con = conectar(ruta_db)
    try:
        for hoja in hojas:
            registros = leer_destinatarios_hoja(excel_path, hoja)
            if registros is None:
                log(
                    f"[AVISO] No se encontró la columna 'Nombre' en '{hoja}'. Pestaña omitida."
                )
                resultado["hojas_con_error"].append(hoja)
                continue

            errores = _validar_registros(registros)
            if errores:
                log(f"[ERROR] Pestaña '{hoja}' NO cargada:")
                for e in errores:
                    log(f"   - {e}")
                resultado["hojas_con_error"].append(hoja)
                continue

            nuevos_hoja = 0
            for reg in registros:
                nombre_norm = normalizar_texto(reg["nombre"])
                if not nombre_norm:
                    continue

                existe = con.execute(
                    "SELECT id FROM destinatarios"
                    " WHERE excel_origen=? AND pestana=? AND id_excel=?",
                    (excel, hoja, reg["id"]),
                ).fetchone()
                if existe:
                    resultado["ya_existian"] += 1
                    continue

                otro = con.execute(
                    "SELECT excel_origen, pestana, id_excel FROM destinatarios"
                    " WHERE nombre_norm=?",
                    (nombre_norm,),
                ).fetchone()
                if otro:
                    resultado["conflictos"] += 1
                    log(
                        f"   [AVISO] '{reg['nombre']}' (ID {reg['id']}, '{hoja}') ya existe en"
                        f" {otro['excel_origen']} / {otro['pestana']} / ID {otro['id_excel']}."
                        " No se insertó: revisa los Excel."
                    )
                    continue

                con.execute(
                    "INSERT INTO destinatarios (excel_origen, pestana, id_excel, nombre,"
                    " nombre_norm, cargo, entidad, correo) VALUES (?,?,?,?,?,?,?,?)",
                    (
                        excel,
                        hoja,
                        reg["id"],
                        reg["nombre"],
                        nombre_norm,
                        reg["cargo"],
                        reg["entidad"],
                        reg["correo"],
                    ),
                )
                nuevos_hoja += 1
            con.commit()
            resultado["nuevos"] += nuevos_hoja
            log(f"  - '{hoja}': {nuevos_hoja} destinatarios nuevos")
    finally:
        con.close()
    return resultado


# ----------------------------------------------------------------------
# Recarga de emergencia (upsert con vista previa)
# ----------------------------------------------------------------------
def planificar_recarga(excel_path, hojas, ruta_db=RUTA_DB):
    """
    Compara el Excel con la BD SIN escribir nada. Devuelve un 'plan':
      nuevos, actualizar, sin_cambio, ausentes, conflictos, hojas_con_error,
      archivo_desconocido (True si ese Excel no está en la BD: posible renombrado).
    """
    excel = os.path.basename(excel_path)
    plan = {
        "excel": excel,
        "nuevos": [],
        "actualizar": [],
        "sin_cambio": 0,
        "ausentes": [],
        "conflictos": [],
        "hojas_con_error": [],
        "archivo_desconocido": False,
    }
    vistos_nuevos = set()
    con = conectar(ruta_db)
    try:
        total_bd = con.execute("SELECT COUNT(*) FROM destinatarios").fetchone()[0]
        en_archivo = con.execute(
            "SELECT COUNT(*) FROM destinatarios WHERE excel_origen=?", (excel,)
        ).fetchone()[0]
        plan["archivo_desconocido"] = total_bd > 0 and en_archivo == 0

        for hoja in hojas:
            registros = leer_destinatarios_hoja(excel_path, hoja)
            if registros is None:
                plan["hojas_con_error"].append(f"{hoja}: sin columna Nombre")
                continue
            errores = _validar_registros(registros)
            if errores:
                plan["hojas_con_error"].extend(f"{hoja}: {e}" for e in errores)
                continue

            ids_excel = set()
            for reg in registros:
                nn = normalizar_texto(reg["nombre"])
                if not nn:
                    continue
                ids_excel.add(reg["id"])
                fila = con.execute(
                    "SELECT * FROM destinatarios"
                    " WHERE excel_origen=? AND pestana=? AND id_excel=?",
                    (excel, hoja, reg["id"]),
                ).fetchone()

                if fila is None:
                    otro = con.execute(
                        "SELECT id FROM destinatarios WHERE nombre_norm=?", (nn,)
                    ).fetchone()
                    if otro or nn in vistos_nuevos:
                        plan["conflictos"].append(
                            f"{reg['nombre']} (ID {reg['id']}, '{hoja}'): nombre ya existe"
                        )
                    else:
                        vistos_nuevos.add(nn)
                        plan["nuevos"].append(
                            {**reg, "pestana": hoja, "nombre_norm": nn}
                        )
                    continue

                cambios = {
                    c: reg[c] for c in CAMPOS_EDITABLES if (fila[c] or "") != reg[c]
                }
                if not cambios:
                    plan["sin_cambio"] += 1
                    continue
                if "nombre" in cambios:
                    otro = con.execute(
                        "SELECT id FROM destinatarios WHERE nombre_norm=? AND id<>?",
                        (nn, fila["id"]),
                    ).fetchone()
                    if otro:
                        plan["conflictos"].append(
                            f"{reg['nombre']} (ID {reg['id']}, '{hoja}'): el nuevo nombre ya existe"
                        )
                        continue
                plan["actualizar"].append(
                    {
                        "db_id": fila["id"],
                        "nombre_actual": fila["nombre"],
                        "pestana": hoja,
                        "cambios": cambios,
                        "nombre_norm": nn,
                    }
                )

            en_bd = con.execute(
                "SELECT id_excel, nombre FROM destinatarios WHERE excel_origen=? AND pestana=?",
                (excel, hoja),
            ).fetchall()
            for a in en_bd:
                if a["id_excel"] not in ids_excel:
                    plan["ausentes"].append(
                        f"{a['nombre']} (ID {a['id_excel']}, '{hoja}')"
                    )
    finally:
        con.close()
    return plan


def aplicar_recarga(plan, ruta_db=RUTA_DB, log=print):
    """
    Aplica un plan de planificar_recarga(). NUNCA borra destinatarios y
    NUNCA toca estado, archivo ni lote de los existentes.
    """
    excel = plan["excel"]
    insertados = actualizados = 0
    con = conectar(ruta_db)
    try:
        for reg in plan["nuevos"]:
            try:
                con.execute(
                    "INSERT INTO destinatarios (excel_origen, pestana, id_excel, nombre,"
                    " nombre_norm, cargo, entidad, correo) VALUES (?,?,?,?,?,?,?,?)",
                    (
                        excel,
                        reg["pestana"],
                        reg["id"],
                        reg["nombre"],
                        reg["nombre_norm"],
                        reg["cargo"],
                        reg["entidad"],
                        reg["correo"],
                    ),
                )
                insertados += 1
            except sqlite3.IntegrityError as e:
                log(f"[AVISO] No se pudo insertar '{reg['nombre']}': {e}")

        for item in plan["actualizar"]:
            cambios = dict(item["cambios"])
            if "nombre" in cambios:
                cambios["nombre_norm"] = item["nombre_norm"]
            sets = ", ".join(
                f"{campo}=?" for campo in cambios
            )  # campos fijos, no del usuario
            try:
                con.execute(
                    f"UPDATE destinatarios SET {sets} WHERE id=?",
                    (*cambios.values(), item["db_id"]),
                )
                actualizados += 1
            except sqlite3.IntegrityError as e:
                log(f"[AVISO] No se pudo actualizar '{item['nombre_actual']}': {e}")
        con.commit()
    finally:
        con.close()
    log(f"[OK] Recarga aplicada: {insertados} nuevos, {actualizados} actualizados.")
    return {"insertados": insertados, "actualizados": actualizados}


# ----------------------------------------------------------------------
# Cruce de nombres
# ----------------------------------------------------------------------
def extraer_zona_destinatario(texto_norm):
    """Devuelve la zona donde debería estar el nombre del destinatario."""
    m = re.search(rf"\b{TRATAMIENTOS}\b", texto_norm)
    if m:
        return texto_norm[m.end() : m.end() + 200]
    return texto_norm[:500]  # Plan B: encabezado de la página


def _puntuar(nombre_norm, zona):
    if re.search(rf"\b{re.escape(nombre_norm)}\b", zona):
        return 100
    # partial_ratio compara la cadena MÁS CORTA contra trozos de la larga. Si la zona
    # es más corta que el nombre, un fragmento ("carlos" o hasta una letra suelta)
    # puntuaría 100. En ese caso se compara la cadena completa, que da puntaje bajo.
    if len(zona) < len(nombre_norm) * FACTOR_LONGITUD:
        return fuzz.ratio(nombre_norm, zona)
    return fuzz.partial_ratio(nombre_norm, zona)


def buscar_destinatario(texto_norm, candidatos):
    """
    candidatos: lista de tuplas (id, nombre_norm, estado), con TODOS los
    destinatarios (pendientes y procesados).

    Devuelve (id, estado, puntaje, ambiguo). Si no hay match: (None, None, mejor_puntaje, False).

    Se puntúa a todos juntos. Así, si la carta pertenece a alguien ya procesado,
    gana él (duplicado) en lugar de que se la asignen a un pendiente parecido.
    En empates de puntaje gana el nombre más largo (más específico).
    """
    if len(texto_norm) < MIN_CARACTERES_OCR:  # página en blanco, mancha o reverso
        return None, None, 0, False

    ranking = []
    for zona, umbral in (
        (extraer_zona_destinatario(texto_norm), UMBRAL_ZONA),
        (texto_norm, UMBRAL_PAGINA),  # respaldo: todo el texto leído, más estricto
    ):
        ranking = sorted(
            (
                (_puntuar(nn, zona), len(nn), id_, estado)
                for id_, nn, estado in candidatos
                if nn
            ),
            reverse=True,
        )
        if ranking and ranking[0][0] >= umbral:
            ambiguo = len(ranking) > 1 and ranking[0][0] - ranking[1][0] < DIF_AMBIGUO
            return ranking[0][2], ranking[0][3], ranking[0][0], ambiguo

    mejor = ranking[0][0] if ranking else 0
    return None, None, mejor, False


# ----------------------------------------------------------------------
# Procesamiento de lotes
# ----------------------------------------------------------------------
def _escribir_pagina(lector, indice, ruta_destino):
    escritor = PdfWriter()
    escritor.add_page(lector.pages[indice])
    with open(ruta_destino, "wb") as salida:
        escritor.write(salida)


def _ocr_pagina(pdf_path, num_pagina):
    """Rasteriza SOLO una página del PDF y le hace OCR a su parte superior."""
    kwargs = {"dpi": DPI_OCR, "first_page": num_pagina, "last_page": num_pagina}
    if RUTA_POPPLER:
        kwargs["poppler_path"] = RUTA_POPPLER
    imagen = convert_from_path(pdf_path, **kwargs)[0]
    ancho, alto = imagen.size
    zona_sup = imagen.crop((0, 0, ancho, int(alto * FRACCION_SUPERIOR)))
    return pytesseract.image_to_string(zona_sup, lang="spa")


def procesar_lote(
    pdf_path,
    ruta_db=RUTA_DB,
    carpeta_salida=CARPETA_SALIDA,
    log=print,
    progreso=None,
    detener=None,
):
    """
    Procesa un PDF de N cartas contra la lista maestra guardada en la BD.

    - Si el mismo PDF (por huella SHA-256) ya se procesó por completo, no hace nada.
    - Si quedó a medias (corte de luz, cancelado) o con páginas en ERROR, REANUDA
      el mismo lote: salta las páginas ya hechas y reintenta las que dieron ERROR.
    - Commit tras CADA página.

    Estados de página: OK, REVISAR (ambiguo), DUPLICADO_REVISAR, SIN_MATCH, ERROR.
    Devuelve un dict con el resumen de esta ejecución (o None si no hizo nada).
    """
    inicio = time.time()
    os.makedirs(carpeta_salida, exist_ok=True)
    con = conectar(ruta_db)

    try:
        filas = con.execute(
            "SELECT id, nombre, nombre_norm, estado FROM destinatarios"
        ).fetchall()
        if not filas:
            log("[ERROR] La lista maestra está vacía. Carga primero el Excel.")
            return None

        nombres = {f["id"]: f["nombre"] for f in filas}
        normalizados = {f["id"]: f["nombre_norm"] for f in filas}
        estados = {f["id"]: f["estado"] for f in filas}

        hash_pdf = calcular_hash(pdf_path)
        lector = PdfReader(pdf_path)
        total = len(lector.pages)

        previo = con.execute(
            "SELECT id, estado, duracion_seg FROM lotes WHERE hash_pdf=? ORDER BY id DESC LIMIT 1",
            (hash_pdf,),
        ).fetchone()

        previas = {}
        duracion_previa = 0.0
        if previo:
            lote_id = previo["id"]
            previas = {
                r["num_pagina"]: r["estado"]
                for r in con.execute(
                    "SELECT num_pagina, estado FROM paginas WHERE lote_id=?", (lote_id,)
                )
            }
            con_error = sum(1 for e in previas.values() if e == "ERROR")
            if previo["estado"] == "COMPLETO" and con_error == 0:
                log(
                    f"[AVISO] Este PDF ya se procesó por completo en el lote {lote_id}."
                    " No se hace nada."
                )
                return None
            hechas = sum(1 for e in previas.values() if e != "ERROR")
            log(
                f"[REANUDAR] Este PDF corresponde al lote {lote_id}: {hechas} páginas ya"
                f" procesadas, {con_error} con ERROR se reintentarán."
            )
            duracion_previa = previo["duracion_seg"] or 0.0
            con.execute("UPDATE lotes SET estado='EN_PROCESO' WHERE id=?", (lote_id,))
            con.commit()
        else:
            lote_id = con.execute(
                "INSERT INTO lotes (fecha, pdf_origen, hash_pdf, total_paginas)"
                " VALUES (?, ?, ?, ?)",
                (
                    datetime.now().isoformat(timespec="seconds"),
                    pdf_path,
                    hash_pdf,
                    total,
                ),
            ).lastrowid
            con.commit()

        pendientes_antes = sum(1 for e in estados.values() if e == "Pendiente")
        log(
            f"\nLote {lote_id}: {total} páginas en '{pdf_path}'."
            f" Destinatarios pendientes: {pendientes_antes}\n"
        )

        cont = {
            "OK": 0,
            "REVISAR": 0,
            "DUPLICADO_REVISAR": 0,
            "SIN_MATCH": 0,
            "ERROR": 0,
        }
        cancelado = False

        for num in range(1, total + 1):
            if detener is not None and detener.is_set():
                log("\n[CANCELADO] Proceso detenido por el usuario.")
                cancelado = True
                break

            if num in previas and previas[num] != "ERROR":
                if progreso is not None:
                    progreso(num, total)
                continue

            try:
                texto = _ocr_pagina(pdf_path, num)
                texto_norm = normalizar_texto(texto)

                candidatos = [(i, normalizados[i], estados[i]) for i in normalizados]
                dest_id, estado_dest, puntaje, ambiguo = buscar_destinatario(
                    texto_norm, candidatos
                )

                if dest_id is None:
                    estado_pag = "SIN_MATCH"
                    base = f"SIN_MATCH_lote{lote_id}_pagina_{num}"
                elif estado_dest == "Pendiente":
                    estado_pag = "REVISAR" if ambiguo else "OK"
                    base = nombre_archivo_seguro(nombres[dest_id])
                else:  # ya estaba procesado: segunda carta para la misma persona
                    estado_pag = "DUPLICADO_REVISAR"
                    base = f"DUPLICADO_{nombre_archivo_seguro(nombres[dest_id])}"

                ruta = ruta_disponible(carpeta_salida, base)
                _escribir_pagina(lector, num - 1, ruta)
                archivo = os.path.basename(ruta)

                con.execute(
                    "DELETE FROM paginas WHERE lote_id=? AND num_pagina=? AND estado='ERROR'",
                    (lote_id, num),
                )
                con.execute(
                    "INSERT INTO paginas (lote_id, num_pagina, destinatario_id,"
                    " puntaje, estado, archivo, texto_ocr) VALUES (?,?,?,?,?,?,?)",
                    (lote_id, num, dest_id, puntaje, estado_pag, archivo, texto_norm),
                )
                if estado_pag in ("OK", "REVISAR"):
                    con.execute(
                        "UPDATE destinatarios SET estado='Procesado', archivo=?,"
                        " lote_id=? WHERE id=?",
                        (archivo, lote_id, dest_id),
                    )
                    estados[dest_id] = "Procesado"
                con.commit()  # <- guardado página por página

                cont[estado_pag] += 1
                prefijo = f"[{num}/{total}]"
                if estado_pag == "OK":
                    log(f"{prefijo} OK ({puntaje:.0f}%) -> {archivo}")
                elif estado_pag == "REVISAR":
                    log(
                        f"{prefijo} REVISAR (posible empate, {puntaje:.0f}%) -> {archivo}"
                    )
                elif estado_pag == "DUPLICADO_REVISAR":
                    log(
                        f"{prefijo} DUPLICADO: ya existía carta de esta persona -> {archivo}"
                    )
                else:
                    log(f"{prefijo} SIN_MATCH (mejor: {puntaje:.0f}%) -> {archivo}")
                    if len(texto_norm) < MIN_CARACTERES_OCR:
                        log("   (página casi vacía: ¿hoja en blanco o reverso?)")
                    log(f"   OCR leído: {texto_norm[:200]}")

            except Exception as e:
                con.rollback()
                con.execute(
                    "DELETE FROM paginas WHERE lote_id=? AND num_pagina=? AND estado='ERROR'",
                    (lote_id, num),
                )
                con.execute(
                    "INSERT INTO paginas (lote_id, num_pagina, estado, texto_ocr)"
                    " VALUES (?, ?, 'ERROR', ?)",
                    (lote_id, num, str(e)),
                )
                con.commit()
                cont["ERROR"] += 1
                log(f"[{num}/{total}] ERROR procesando página: {e}")

            if progreso is not None:
                progreso(num, total)

        segundos = time.time() - inicio
        con.execute(
            "UPDATE lotes SET estado=?, duracion_seg=? WHERE id=?",
            (
                "CANCELADO" if cancelado else "COMPLETO",
                duracion_previa + segundos,
                lote_id,
            ),
        )
        con.commit()
        pendientes_despues = con.execute(
            "SELECT COUNT(*) FROM destinatarios WHERE estado='Pendiente'"
        ).fetchone()[0]
    finally:
        con.close()

    resumen = {
        "lote_id": lote_id,
        "total_paginas": total,
        "cancelado": cancelado,
        "pendientes_despues": pendientes_despues,
        "segundos": segundos,
        **cont,
    }

    log("\n" + "=" * 50)
    log("¡LOTE FINALIZADO!" if not cancelado else "LOTE INTERRUMPIDO")
    log(f"📄 Páginas del PDF: {total}")
    log(f"✅ OK: {cont['OK']}")
    log(f"🔎 Para REVISAR (empate): {cont['REVISAR']}")
    log(f"♻️ Duplicados para revisar: {cont['DUPLICADO_REVISAR']}")
    log(f"⚠️ Sin coincidencia: {cont['SIN_MATCH']}")
    log(f"❌ Con error: {cont['ERROR']}")
    log(f"👥 Destinatarios aún pendientes: {pendientes_despues}")
    log(
        f"⏱️ Tiempo de esta ejecución: {int(segundos // 60)} min {int(segundos % 60)} s"
    )
    log("=" * 50)
    return resumen


def procesar_carpeta(
    carpeta,
    ruta_db=RUTA_DB,
    carpeta_salida=CARPETA_SALIDA,
    log=print,
    progreso=None,
    detener=None,
):
    """Procesa, en orden alfabético, todos los PDFs de una carpeta (un lote por PDF)."""
    pdfs = sorted(f for f in os.listdir(carpeta) if f.lower().endswith(".pdf"))
    if not pdfs:
        log(f"[AVISO] No se encontraron PDFs en '{carpeta}'.")
        return []

    log(f"Se procesarán {len(pdfs)} PDF(s) de '{carpeta}'.")
    resumenes = []
    for i, nombre in enumerate(pdfs, 1):
        if detener is not None and detener.is_set():
            log("[CANCELADO] Se detuvo el procesamiento de la carpeta.")
            break
        log(f"\n######## PDF {i}/{len(pdfs)}: {nombre} ########")
        try:
            r = procesar_lote(
                os.path.join(carpeta, nombre),
                ruta_db,
                carpeta_salida,
                log,
                progreso,
                detener,
            )
            if r:
                resumenes.append(r)
        except Exception as e:
            log(f"[ERROR] No se pudo procesar '{nombre}': {e}")
    return resumenes


# ----------------------------------------------------------------------
# Verificación de completitud
# ----------------------------------------------------------------------
def verificar_completitud(ruta_db=RUTA_DB):
    """
    Compara páginas procesadas (en cualquier estado) con el tamaño de la lista.
      procesadas < lista  -> EN CURSO
      procesadas = lista  -> COMPLETO
      procesadas > lista  -> INCONSISTENTE
    Además cuenta páginas por revisar y destinatarios sin carta.
    """
    con = conectar(ruta_db)
    try:

        def uno(sql):
            return con.execute(sql).fetchone()[0]

        total_lista = uno("SELECT COUNT(*) FROM destinatarios")
        procesadas = uno("SELECT COUNT(*) FROM paginas")
        por_revisar = uno("SELECT COUNT(*) FROM paginas WHERE estado <> 'OK'")
        pendientes = uno("SELECT COUNT(*) FROM destinatarios WHERE estado='Pendiente'")
    finally:
        con.close()

    if procesadas < total_lista:
        estado = "EN CURSO"
    elif procesadas == total_lista:
        estado = "COMPLETO"
    else:
        estado = "INCONSISTENTE"

    return {
        "total_lista": total_lista,
        "paginas_procesadas": procesadas,
        "diferencia": procesadas - total_lista,
        "estado": estado,
        "paginas_por_revisar": por_revisar,
        "destinatarios_pendientes": pendientes,
    }


def mostrar_verificacion(v, log=print):
    log("\n" + "-" * 50)
    log("VERIFICACIÓN")
    log(
        f"Lista maestra: {v['total_lista']} | Páginas procesadas: {v['paginas_procesadas']}"
    )
    log(f"Estado del proceso: {v['estado']}")

    if v["estado"] == "EN CURSO":
        log(f"Faltan {-v['diferencia']} carta(s) por procesar.")
        log(f"Páginas por revisar hasta ahora: {v['paginas_por_revisar']}")
        log(f"Destinatarios sin carta aún: {v['destinatarios_pendientes']}")
    else:
        if v["estado"] == "INCONSISTENTE":
            log(
                f"⚠️ Hay {v['diferencia']} página(s) de más que destinatarios."
                " ¿Se repitió un PDF o hay cartas de más?"
            )
        avisos = []
        if v["paginas_por_revisar"]:
            avisos.append(
                f"⚠️ WARNING: {v['paginas_por_revisar']} página(s) requieren revisión manual"
                " (SIN_MATCH, REVISAR, DUPLICADO_REVISAR o ERROR)."
            )
        if v["destinatarios_pendientes"]:
            avisos.append(
                f"⚠️ WARNING: {v['destinatarios_pendientes']} destinatario(s) siguen sin carta asignada."
            )
        if avisos:
            for a in avisos:
                log(a)
        elif v["estado"] == "COMPLETO":
            log(
                "✅ Todo en orden: cada destinatario tiene su carta y no hay nada por revisar."
            )
    log("-" * 50)


# ----------------------------------------------------------------------
# Excel de control (único, generado desde la BD)
# ----------------------------------------------------------------------
def _ajustar_hoja(ws, filtro=True):
    if filtro:
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"
    for col in ws.columns:
        largo = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[col[0].column_letter].width = min(largo + 2, 60)


def _guardar_excel_control(
    ruta, indicadores, por_estado, por_lote, destinatarios, paginas, por_revisar
):
    with pd.ExcelWriter(ruta, engine="openpyxl") as escritor:
        # --- Pestaña Resumen: tres bloques apilados ---
        fila = 1
        for titulo, df in (
            ("INDICADORES", indicadores),
            ("PÁGINAS POR ESTADO", por_estado),
            ("DETALLE POR LOTE", por_lote),
        ):
            df.to_excel(escritor, sheet_name="Resumen", startrow=fila, index=False)
            ws = escritor.sheets["Resumen"]
            ws.cell(row=fila, column=1, value=titulo).font = Font(bold=True)
            fila += len(df) + 3
        _ajustar_hoja(escritor.sheets["Resumen"], filtro=False)

        # --- Resto de pestañas ---
        for nombre, df in (
            ("Destinatarios", destinatarios),
            ("Páginas", paginas),
            ("Por revisar", por_revisar),
        ):
            df.to_excel(escritor, sheet_name=nombre, index=False)
            _ajustar_hoja(escritor.sheets[nombre])


def generar_reporte_control(
    ruta_db=RUTA_DB,
    carpeta_reportes=CARPETA_REPORTES,
    nombre="reporte_control.xlsx",
    log=print,
):
    """
    Genera el Excel de control con las pestañas Resumen, Destinatarios,
    Páginas y Por revisar. Es una vista de la BD: se puede regenerar siempre.
    Devuelve la ruta, o None si el archivo está abierto en Excel.
    """
    os.makedirs(carpeta_reportes, exist_ok=True)
    v = verificar_completitud(ruta_db)

    con = conectar(ruta_db)
    try:
        destinatarios = pd.read_sql_query(
            "SELECT id_excel AS ID, nombre AS Nombre, cargo AS Cargo, entidad AS Entidad,"
            " correo AS Correo, excel_origen AS Excel, pestana AS Pestaña, estado AS Estado,"
            " archivo AS Archivo_PDF, lote_id AS Lote FROM destinatarios"
            " ORDER BY excel_origen, pestana, CAST(id_excel AS INTEGER), id_excel",
            con,
        )
        paginas = pd.read_sql_query(
            "SELECT p.lote_id AS Lote, p.num_pagina AS Pagina, d.nombre AS Destinatario,"
            " d.cargo AS Cargo, d.entidad AS Entidad, p.puntaje AS Puntaje,"
            " p.estado AS Estado, p.archivo AS Archivo, p.texto_ocr AS Texto_OCR"
            " FROM paginas p LEFT JOIN destinatarios d ON d.id = p.destinatario_id"
            " ORDER BY p.lote_id, p.num_pagina",
            con,
        )
        por_estado = pd.read_sql_query(
            "SELECT estado AS Estado, COUNT(*) AS Paginas FROM paginas GROUP BY estado",
            con,
        )
        por_lote = pd.read_sql_query(
            "SELECT l.id AS Lote, l.fecha AS Fecha, l.pdf_origen AS PDF,"
            " l.total_paginas AS Paginas, l.estado AS Estado_lote,"
            " COALESCE(SUM(p.estado='OK'),0) AS OK,"
            " COALESCE(SUM(p.estado='REVISAR'),0) AS REVISAR,"
            " COALESCE(SUM(p.estado='DUPLICADO_REVISAR'),0) AS DUPLICADO_REVISAR,"
            " COALESCE(SUM(p.estado='SIN_MATCH'),0) AS SIN_MATCH,"
            " COALESCE(SUM(p.estado='ERROR'),0) AS ERROR,"
            " ROUND(l.duracion_seg,0) AS Duracion_seg"
            " FROM lotes l LEFT JOIN paginas p ON p.lote_id = l.id"
            " GROUP BY l.id ORDER BY l.id",
            con,
        )
    finally:
        con.close()

    por_revisar = paginas[paginas["Estado"] != "OK"].copy()
    por_revisar["Resuelto (Sí/No)"] = ""
    por_revisar["Observación"] = ""

    indicadores = pd.DataFrame(
        [
            ("Generado el", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ("Estado del proceso", v["estado"]),
            ("Destinatarios en la lista maestra", v["total_lista"]),
            ("Páginas procesadas (cualquier estado)", v["paginas_procesadas"]),
            ("Diferencia (procesadas - lista)", v["diferencia"]),
            ("Páginas por revisar manualmente", v["paginas_por_revisar"]),
            ("Destinatarios sin carta asignada", v["destinatarios_pendientes"]),
        ],
        columns=["Indicador", "Valor"],
    )

    ruta = os.path.join(carpeta_reportes, nombre)
    try:
        _guardar_excel_control(
            ruta, indicadores, por_estado, por_lote, destinatarios, paginas, por_revisar
        )
    except PermissionError:
        log(
            f"[ERROR] No se pudo guardar '{ruta}'. Si está abierto en Excel, ciérralo y reintenta."
        )
        return None

    log(f"[REPORTE] Excel de control -> {ruta}")
    mostrar_verificacion(v, log)
    return ruta


# ----------------------------------------------------------------------
# Reinicio (fin de campaña)
# ----------------------------------------------------------------------
def reiniciar_campana(
    ruta_db=RUTA_DB,
    carpeta_salida=CARPETA_SALIDA,
    carpeta_reportes=CARPETA_REPORTES,
    borrar_pdfs=True,
    log=print,
):
    """
    Cierra la campaña: genera un reporte de control FINAL, hace un respaldo de
    la BD, y vacía TODO (incluida la lista de destinatarios).
    Si no se puede generar el reporte final, NO borra nada.
    """
    if not os.path.exists(ruta_db):
        log("[AVISO] No hay base de datos que reiniciar.")
        return False

    sello = datetime.now().strftime("%Y%m%d_%H%M%S")

    if verificar_completitud(ruta_db)["total_lista"] > 0:
        reporte = generar_reporte_control(
            ruta_db, carpeta_reportes, f"reporte_control_FINAL_{sello}.xlsx", log
        )
        if reporte is None:
            log("[ERROR] No se pudo generar el reporte final. Reinicio cancelado.")
            return False

    respaldo = f"{os.path.splitext(ruta_db)[0]}_respaldo_{sello}.db"
    origen = sqlite3.connect(ruta_db)
    destino = sqlite3.connect(respaldo)
    origen.backup(destino)  # copia consistente de la BD
    destino.close()
    origen.close()
    log(f"[RESPALDO] {respaldo}")

    con = conectar(ruta_db)
    try:
        con.execute("DELETE FROM paginas")
        con.execute("DELETE FROM lotes")
        con.execute("DELETE FROM destinatarios")
        con.execute(
            "DELETE FROM sqlite_sequence WHERE name IN ('paginas','lotes','destinatarios')"
        )
        con.commit()
    finally:
        con.close()

    if borrar_pdfs and os.path.isdir(carpeta_salida):
        for f in os.listdir(carpeta_salida):
            if f.lower().endswith(".pdf"):
                os.remove(os.path.join(carpeta_salida, f))
        log(f"[OK] PDFs de '{carpeta_salida}' eliminados.")

    log("[OK] Campaña reiniciada. La base de datos quedó vacía.")
    return True
