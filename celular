# **Importar Librerias**
"""

!pip install pytesseract pdfplumber opencv-python pandas
!pip install paddleocr
!pip install paddlepaddle
!pip install pdfplumber opencv-python pandas
!apt-get update
!apt-get install -y tesseract-ocr
!pip install pytesseract
!pip install pandas openpyxl
!pip install dateparser

import os
import cv2
import pytesseract
import pandas as pd
from google.colab import drive
import numpy as np
import pdfplumber
from paddleocr import PaddleOCR
import re

"""# **Conectar a Drive**"""

# 📌 Paso 2: Montar Google Drive
drive.mount('/content/drive')

# 📌 Paso 3: Definir rutas de archivos
folder = "/content/drive/MyDrive/Automatizaciones/Observaciones/"

"""# **Análisis Observaciones**"""

# 📌 Paso 3: Definir carpeta y buscar primer archivo Excel
folder = "/content/drive/MyDrive/Automatizaciones/Observaciones/"

# Buscar el primer archivo que termine en .xlsx o .xls
archivo_excel = next((f for f in os.listdir(folder) if f.endswith(('.xlsx', '.xls'))), None)

if archivo_excel:
    ruta_archivo = os.path.join(folder, archivo_excel)
    print(f"Archivo encontrado: {ruta_archivo}")
else:
    raise FileNotFoundError("No se encontró ningún archivo Excel en la carpeta.")

ruta_archivo

# Paso 5: Leer el Excel
df = pd.read_excel(ruta_archivo)

# Verifica que la columna AB (índice 27) existe
if len(df.columns) < 28:
    raise IndexError("El archivo no tiene suficientes columnas. ¿Estás seguro que hay una columna AB?")

# Paso 6: Extraer celulares desde la columna AB (índice 27)
columna_observaciones = df.columns[27]  # Columna AB
print(f"📌 Usando columna: {columna_observaciones}")

regex_celular = r"(3\d{2}[ -.]?\d{3}[ -.]?\d{4})"

df['celular_extraído'] = df[columna_observaciones].astype(str).apply(
    lambda texto: re.findall(regex_celular, texto)[0] if re.findall(regex_celular, texto) else None
)

# Paso 7: Mostrar primeros resultados
df[[columna_observaciones, 'celular_extraído']].head()

from google.colab import files

# Paso 1: Seleccionar las columnas deseadas
columna_A = df.columns[0]
columna_D = df.columns[3]
columna_I = df.columns[8]
columna_AB = df.columns[27]

# Paso 2: Crear nuevo DataFrame
df_exportar = df[[columna_A, columna_D, columna_I, columna_AB, 'celular_extraído']].copy()
df_exportar.columns = ['ID', 'Nombre', 'Ciudad', 'Observación', 'Celular']

# Paso 3: Guardar archivo temporal en Colab
nombre_archivo = "resumen_contactos.xlsx"
df_exportar.to_excel(nombre_archivo, index=False)

# Paso 4: Descargar archivo
files.download(nombre_archivo)

"""# **Extraer Fecha**"""

import dateparser
import re

# Lista de etiquetas válidas (en minúsculas)
etiquetas_validas = [
    'fecha asignada',
    'fecha de visita',
    'fecha visita',
    'fecha programada',
    'programada para',
    'agendada para',
    'agenda',
    'día de la visita'
]

# Patrones de fechas flexibles
patrones_fecha = [
    r'\d{4}-\d{2}-\d{2}',                         # 2025-07-21
    r'\d{1,2}/\d{1,2}/\d{2,4}',                   # 21/07/2025
    r'\d{1,2}\s+de\s+[a-záéíóúñ]+(?:\s+\d{4})?',   # 21 de julio [de 2025]
    r'[a-záéíóúñ]+\s+\d{1,2}(?:,\s*\d{4})?',       # julio 21
    r'\b\d{1,2}\b\s+[a-záéíóúñ]+'                 # 21 julio
]

# Función para extraer fecha desde texto, solo si está cerca de etiquetas válidas
def extraer_fecha_relevante(texto):
    texto = str(texto).lower()
    for etiqueta in etiquetas_validas:
        if etiqueta in texto:
            # Capturar texto cercano a la etiqueta
            seccion = texto.split(etiqueta, 1)[-1][:50]
            for patron in patrones_fecha:
                match = re.search(patron, seccion)
                if match:
                    fecha = dateparser.parse(match.group(0), languages=['es'])
                    if fecha:
                        return fecha.date()
    return None

# Aplicar sobre columna de observaciones
df['fecha_solicitada'] = df[columna_AB].apply(extraer_fecha_relevante)

df['fecha_solicitada']

sin_fecha = df['fecha_solicitada'].isna().sum()
print(f"📌 Registros sin fecha solicitada: {sin_fecha}")
