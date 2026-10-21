# 📄 Sistema de División y Renombrado Automatizado de PDFs mediante OCR

Herramienta de escritorio desarrollada para automatizar el flujo de procesamiento masivo de cartas y comunicados oficiales. El sistema toma un "PDF Maestro" unificado (multi-página), lo divide en documentos individuales, realiza reconocimiento óptico de caracteres (**OCR**) para identificar al destinatario en el texto de cada hoja y lo renombra cruzando la información con una base de datos en Excel.

## 💡 Motivación y Problema de Negocio

En dependencias institucionales y despachos gubernamentales, es habitual recibir o generar un solo documento PDF que agrupa cientos de cartas o invitaciones individualizadas.

El procesamiento manual de este tipo de archivos requiere:

1. Abrir un visor de PDF.

2. Identificar visualmente el nombre de la persona en la carta.

3. Extraer esa hoja específica.

4. Guardarla y renombrarla manualmente con el nombre exacto del destinatario.

Este flujo manual genera **cuellos de botella operacionales**, errores de digitación y costos considerables en horas-hombre. Esta solución reduce el tiempo de procesamiento de horas a cuestión de minutos, garantizando una auditoría precisa e integridad en los archivos generados.

## 🛠️ Arquitectura del Sistema y Tecnologías

El proyecto fue diseñado bajo el principio de **separación de responsabilidades (SoC)**, dividiendo estrictamente la lógica de procesamiento pesado de la interfaz gráfica de usuario.

```
DivisionPDFs/
│
├── venv/                       # Entorno virtual aislado de dependencias
├── pdfs_divididos/             # Directorio de salida para los PDFs procesados
│
├── src/                        # Código fuente modular
│   ├── __init__.py             # Identificador de paquete Python
│   ├── procesador.py           # Motor de OCR, lectura de Excel, división PDF y temporizador
│   └── interfaz.py             # GUI en Tkinter, gestión de Threads y Queue
│
├── BasePruebaDivision.xlsx     # Base de datos local (Destinatarios)
├── InvitacionesCongreMinDir.pdf# Documento de prueba
├── .gitignore                  # Exclusiones de control de versiones
├── main.py                     # Punto de entrada y arranque de la aplicación
└── README.md                   # Documentación principal del proyecto

```

### Componentes Clave:

* **`Threading` & `Queue`**: Manejo de concurrencia. El procesamiento masivo corre en un hilo (*Thread*) secundario para evitar que la interfaz de Tkinter se congele ("No responde"), comunicando progresos en tiempo real mediante una cola de mensajes *thread-safe*.

* **`PyTesseract` + Tesseract-OCR**: Motor de extracción de texto sobre imágenes renderizadas a partir de las páginas PDF.

* **`Pandas` & `OpenPyXL`**: Lectura eficiente de hojas de cálculo dinámicas y extracción de listas maestras.

* **`PyPDF` & `pdf2image`**: Manipulación de estructura de PDFs y conversión de páginas a mapas de bits para su análisis OCR.

* **Normalización de Texto (`unicodedata`)**: Algoritmo de comparación tolerante a fallos que remueve tildes, convierte a minúsculas y elimina espacios duplicados o caracteres especiales fanzine para asegurar un *match* perfecto entre el PDF y el Excel.

## 📋 Requisitos Previos

### 1. Requisitos de Software

* **Python 3.10** o superior.

* **Tesseract-OCR** (instalado en la ruta por defecto: `C:\Program Files\Tesseract-OCR`).

  * **Importante**: Debe incluir el paquete de idioma en español (`spa.traineddata`) dentro del directorio `tessdata`.

* **Poppler para Windows** (necesario para `pdf2image`). Debe estar añadido al `PATH` del sistema.

### 2. Dependencias de Python

Todas las dependencias están aisladas dentro del entorno virtual. Las principales son:

* `pandas`

* `openpyxl`

* `pypdf`

* `pdf2image`

* `pytesseract`

## 🚀 Instalación y Configuración

Sigue estos pasos para desplegar el entorno de desarrollo local:

1. **Clonar el repositorio o navegar a la carpeta del proyecto**:

   ```
   cd C:\Ruta\A\Tu\Proyecto\DivisionPDFs
   
   ```

2. **Crear y activar el entorno virtual (`venv`)**:

   ```
   # En Windows PowerShell
   python -m venv venv
   .\venv\Scripts\Activate.ps1
   
   ```

3. **Instalar dependencias necesarias**:

   ```
   pip install pandas pdf2image pypdf pytesseract openpyxl
   
   ```

4. **Verificar la presencia del archivo de idioma de Tesseract**:
   Asegúrate de que la carpeta `C:\Program Files\Tesseract-OCR\tessdata\` contenga el archivo `spa.traineddata`.

## 💻 Uso de la Aplicación

1. **Ejecutar la aplicación**:
   Con el entorno virtual activo, lanza el punto de entrada principal:

   ```
   python main.py
   
   ```

2. **Flujo de Trabajo en la Interfaz**:

   * **Carga de Datos**: Selecciona el archivo Excel con los destinatarios y el PDF Maestro a procesar.

   * **Ejecución**: Presiona el botón de **Iniciar Procesamiento**.

   * **Monitoreo en Tiempo Real**: Observa el cuadro de consola integrado para ver el avance página por página (estado, destinatario detectado o advertencias de duplicados).

   * **Consulta de Salida**: Al finalizar, los archivos renombrados estarán en la carpeta `pdfs_divididos/` y la consola desplegará el tiempo total de auditoría.

## ⏱️ Auditoría y Métricas

El sistema incluye un módulo interno de cronometraje que registra el tiempo exacto de procesamiento por lote. Esto permite llevar un registro de productividad y evaluar la eficiencia del motor OCR según el volumen de páginas analizadas.

## ✒️ Autoría y Licencia

Desarrollado como una solución de automatización institucional para la optimización de procesos documentales.