crear un pequeño programa de escritorio que tenga como fin poder detectar las facturas en pdf de los mails de provedores de una empresa,
descargarlos automaticamente del outlook y que luego pasen a una carpeta donde seran exportados los datos que se necesita de la factura 
para exportarlos a un excel.
PASOS A REALIZAR:
📥 Descargar automáticamente los PDF desde OneDrive/Outlook.

📄 Procesar el texto del PDF.

🔍 Extraer datos clave.

🧹 Validar y limpiar los datos.

📊 Exportar a Excel.

✅ Mostrar todo con una interfaz simple.

TECNOLOGIAS A USAR :

Python 	Lógica principal y lectura de PDFs
pdfplumber	Leer texto de archivos PDF	
pandas	Manipular datos y generar Excel	
openpyxl	Escribir archivos Excel (.xlsx)	
(opcional luego) PySimpleGUI o Tkinter	Para hacer una interfaz con botón	Ideal para apps de escritorio simples


1 PARTE DEL PROYECTO 08/05/2025
Documentación técnica del código
El código implementa una aplicación de escritorio en Python para procesar facturas en formato PDF y exportar los datos extraídos a un archivo Excel. A continuación, se detalla cómo funciona cada parte del código:

1. Importación de módulos
El código utiliza varias bibliotecas:

os, re, json: Para manejo de archivos, expresiones regulares y configuración.
pdfplumber: Para extraer texto de archivos PDF.
tkinter: Para crear la interfaz gráfica.
openpyxl: Para manipular archivos Excel.
2. Clase FacturadorApp
La clase principal que gestiona la lógica de la aplicación.

2.1 Constructor (__init__)
Configura la ventana principal de la aplicación.
Define los patrones de extracción de datos de las facturas.
Carga la configuración desde un archivo JSON (config.json).
Llama a setup_ui para construir la interfaz gráfica.
2.2 Métodos de configuración
cargar_config: Carga la configuración desde config.json (directorio de salida y último archivo procesado).
guardar_config: Guarda la configuración actual en config.json.
2.3 Interfaz gráfica (setup_ui)
Crea botones, etiquetas y una barra de progreso para interactuar con el usuario.
Incluye un botón principal para iniciar el procesamiento de facturas.
2.4 Procesamiento de facturas
iniciar_proceso: Permite al usuario seleccionar una carpeta de salida y llama a procesar_carpeta.
procesar_carpeta:
Permite seleccionar una carpeta con archivos PDF.
Procesa cada archivo PDF llamando a procesar_archivo_pdf.
Guarda los datos extraídos en un archivo Excel usando guardar_excel.
Muestra un mensaje con el resultado del procesamiento.
2.5 Procesamiento de archivos PDF
procesar_archivo_pdf:

Extrae texto de un archivo PDF usando pdfplumber.
Usa patrones de expresiones regulares para extraer datos como proveedor, CUIT, fecha, número de factura e importe.
Valida que todos los datos requeridos estén presentes.
extraer_dato: Busca un dato específico en el texto del PDF usando el patrón correspondiente.

formatear_importe: Convierte un importe en texto a un formato numérico con separadores correctos.

2.6 Generación de Excel
guardar_excel:
Crea o actualiza un archivo Excel con los datos extraídos.
Ajusta automáticamente el ancho de las columnas.
2.7 Mensajes al usuario
mostrar_resultado: Muestra un mensaje con el resultado del procesamiento y permite al usuario decidir si desea realizar otra operación.
3. Ejecución principal
Si el archivo se ejecuta directamente, se crea una instancia de FacturadorApp y se inicia el bucle principal de la interfaz gráfica
