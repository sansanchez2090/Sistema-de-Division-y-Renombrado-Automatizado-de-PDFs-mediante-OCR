import os
import re
import time
import unicodedata

import pandas as pd
import pytesseract
from pdf2image import convert_from_path
from pypdf import PdfReader, PdfWriter
from rapidfuzz import fuzz

# pip install pandas openpyxl pdf2image pypdf pytesseract rapidfuzz
# Requiere tener instalados Tesseract (con idioma 'spa') y Poppler.

# ----------------------------------------------------------------------
# Configuración
# ----------------------------------------------------------------------
EXCEL_PATH = "BasePruebaDivision.xlsx"
PDF_MAESTRO = "InvitacionesCongreMinDir.pdf"  # El PDF gigante que envía el Despacho
CARPETA_TRABAJO = "./pdfs_divididos"
COLUMNA_NOMBRE = "Nombre"

DPI_OCR = 300  # Resolución para el OCR
FRACCION_SUPERIOR = 0.45  # Se hace OCR solo al 45% superior de la página
UMBRAL_ZONA = 90  # Puntaje mínimo (0-100) buscando en la zona del destinatario
UMBRAL_PAGINA = 96  # Puntaje mínimo buscando en toda la página (más estricto)
# Si Tesseract o Poppler no están en el PATH (Windows), descomenta y ajusta:
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# RUTA_POPPLER = r"C:\poppler\Library\bin"
RUTA_POPPLER = None


# ----------------------------------------------------------------------
# Funciones auxiliares
# ----------------------------------------------------------------------
def normalizar_texto(texto):
    """Minúsculas, sin tildes, sin puntuación y con espacios colapsados."""
    if not isinstance(texto, str):
        return ""
    texto = unicodedata.normalize("NFD", texto.lower())
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = re.sub(r"[^a-z0-9\s]", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def extraer_zona_destinatario(texto_norm):
    """Devuelve la zona donde debería estar el nombre del destinatario."""
    m = re.search(r"\bsenor(?:es|a)?\b", texto_norm)
    if m:
        return texto_norm[m.end() : m.end() + 200]
    return texto_norm[:500]  # Plan B: encabezado de la página


def puntuar(nombre_norm, zona):
    if re.search(rf"\b{re.escape(nombre_norm)}\b", zona):
        return 100
    return fuzz.partial_ratio(nombre_norm, zona)


def buscar_destinatario(texto_norm, lista_nombres):
    """Devuelve (nombre_original, puntaje, ambiguo)."""
    nombres = list(dict.fromkeys(lista_nombres))  # únicos, conserva el orden
    normalizados = {n: normalizar_texto(n) for n in nombres}
    ranking = []

    for zona, umbral in (
        (extraer_zona_destinatario(texto_norm), UMBRAL_ZONA),
        (texto_norm, UMBRAL_PAGINA),  # respaldo: toda la página, más estricto
    ):
        ranking = sorted(
            ((puntuar(nn, zona), n) for n, nn in normalizados.items() if nn),
            reverse=True,
        )
        if ranking and ranking[0][0] >= umbral:
            ambiguo = len(ranking) > 1 and ranking[0][0] - ranking[1][0] < 3
            return ranking[0][1], ranking[0][0], ambiguo

    mejor = ranking[0][0] if ranking else 0
    return None, mejor, False


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


def cargar_nombres_hoja(excel_path, hoja, columna="Nombre", max_filas_busqueda=30):
    """
    Busca en las primeras filas de la hoja la celda cuyo texto sea 'Nombre'
    (sin importar mayúsculas, tildes o espacios) y devuelve los valores de
    esa columna hacia abajo. Devuelve None si no la encuentra.
    """
    crudo = pd.read_excel(excel_path, sheet_name=hoja, header=None, dtype=str)
    objetivo = normalizar_texto(columna)

    for fila in range(min(max_filas_busqueda, len(crudo))):
        valores = [normalizar_texto(v) for v in crudo.iloc[fila]]
        if objetivo in valores:
            col = valores.index(objetivo)
            serie = crudo.iloc[fila + 1 :, col].dropna().astype(str).str.strip()
            return serie[serie != ""].tolist()

    return None


# ----------------------------------------------------------------------
# Inicio
# ----------------------------------------------------------------------
tiempo_inicio_total = time.time()
os.makedirs(CARPETA_TRABAJO, exist_ok=True)

# 1. Selección dinámica de pestañas del Excel
try:
    xl = pd.ExcelFile(EXCEL_PATH)
    todas_las_hojas = xl.sheet_names

    print("¡Hola! Se han detectado las siguientes pestañas en tu Excel:")
    for i, hoja in enumerate(todas_las_hojas):
        print(f"  [{i + 1}] {hoja}")

    seleccion = (
        input(
            "\n¿Qué pestañas deseas procesar? (Escribe los números separados por"
            " comas, o 'todas'): "
        )
        .strip()
        .lower()
    )

    if seleccion == "todas":
        hojas_a_procesar = todas_las_hojas
    else:
        indices = [
            int(x.strip()) - 1 for x in seleccion.split(",") if x.strip().isdigit()
        ]
        hojas_a_procesar = [
            todas_las_hojas[i] for i in indices if 0 <= i < len(todas_las_hojas)
        ]

    # Construir la lista maestra unificada
    lista_maestra = []
    for hoja in hojas_a_procesar:
        nombres = cargar_nombres_hoja(EXCEL_PATH, hoja, COLUMNA_NOMBRE)
        if nombres is None:
            print(
                f"[AVISO] No se encontró la columna '{COLUMNA_NOMBRE}' en las"
                f" primeras 30 filas de la pestaña '{hoja}'."
            )
        else:
            print(f"  - '{hoja}': {len(nombres)} nombres")
            lista_maestra.extend(nombres)

    print(f"\n[OK] Se cargaron {len(lista_maestra)} destinatarios en la lista maestra.")
    if not lista_maestra:
        print(
            "[ERROR CRÍTICO] La lista maestra está vacía. Revisa el nombre de la columna."
        )
        raise SystemExit(1)

except SystemExit:
    raise
except Exception as e:
    print(f"[ERROR CRÍTICO] Leyendo el Excel: {str(e)}")
    raise SystemExit(1)

# 2. Dividir el PDF maestro página por página
print(f"\nDividiendo el PDF maestro '{PDF_MAESTRO}' en páginas individuales...")
lector = PdfReader(PDF_MAESTRO)
total_paginas = len(lector.pages)
archivos_generados = []

for num_pagina in range(total_paginas):
    escritor = PdfWriter()
    escritor.add_page(lector.pages[num_pagina])

    nombre_temporal = f"temp_pagina_{num_pagina + 1}.pdf"
    ruta_temporal = os.path.join(CARPETA_TRABAJO, nombre_temporal)

    with open(ruta_temporal, "wb") as archivo_salida:
        escritor.write(archivo_salida)

    archivos_generados.append(nombre_temporal)

print(
    f"Se han generado {len(archivos_generados)} archivos temporales."
    " Iniciando OCR y renombrado...\n"
)

# 3. OCR + cruce con Excel + renombrado
exitosos = 0
advertencias = 0
errores = 0

for index, archivo in enumerate(archivos_generados):
    ruta_pdf = os.path.join(CARPETA_TRABAJO, archivo)

    try:
        kwargs = {"dpi": DPI_OCR}
        if RUTA_POPPLER:
            kwargs["poppler_path"] = RUTA_POPPLER
        imagenes = convert_from_path(ruta_pdf, **kwargs)

        texto_completo = ""
        for img in imagenes:
            w, h = img.size
            zona_sup = img.crop((0, 0, w, int(h * FRACCION_SUPERIOR)))
            texto_completo += pytesseract.image_to_string(zona_sup, lang="spa") + "\n"

        texto_norm = normalizar_texto(texto_completo)
        destinatario, puntaje, ambiguo = buscar_destinatario(texto_norm, lista_maestra)

        if destinatario:
            base = nombre_archivo_seguro(destinatario)
            nuevo_nombre = ruta_disponible(CARPETA_TRABAJO, base)
            os.rename(ruta_pdf, nuevo_nombre)
            marca = "  << REVISAR (posible empate)" if ambiguo else ""
            print(
                f"[{index + 1}/{total_paginas}] OK ({puntaje:.0f}%) ->"
                f" {os.path.basename(nuevo_nombre)}{marca}"
            )
            exitosos += 1
        else:
            nuevo_nombre = ruta_disponible(
                CARPETA_TRABAJO, f"SIN_MATCH_pagina_{index + 1}"
            )
            os.rename(ruta_pdf, nuevo_nombre)
            print(
                f"[{index + 1}/{total_paginas}] ADVERTENCIA: sin match"
                f" (mejor: {puntaje:.0f}%) -> {os.path.basename(nuevo_nombre)}"
            )
            print("   OCR leído:", texto_norm[:200])
            advertencias += 1

    except Exception as e:
        print(f"[{index + 1}/{total_paginas}] ERROR procesando página: {str(e)}")
        errores += 1

# 4. Resumen
tiempo_total_segundos = time.time() - tiempo_inicio_total
minutos = int(tiempo_total_segundos // 60)
segundos = int(tiempo_total_segundos % 60)

print("\n" + "=" * 50)
print("¡PROCESO FINALIZADO!")
print(f"📄 Total de páginas procesadas: {total_paginas}")
print(f"✅ Archivos renombrados correctamente: {exitosos}")
print(f"⚠️ Páginas sin coincidencia (renombradas SIN_MATCH_*): {advertencias}")
print(f"❌ Páginas con error: {errores}")
print(f"⏱️ Tiempo total de ejecución: {minutos} minutos y {segundos} segundos.")
print("=" * 50)
