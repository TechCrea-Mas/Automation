#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#!/usr/bin/env python
# coding: utf-8

# In[ ]:

import pandas as pd
from pathlib import Path
from datetime import datetime 
import unicodedata
import re



# Crear carpeta de salida si no existe
Path("TEST_salida").mkdir(exist_ok=True)

# Función para normalizar nombres de columna (elimina tildes, pasa a minúscula, reemplaza espacios y signos)
def normalizar(texto):
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("utf-8")
    texto = re.sub(r"[^a-zA-Z0-9]+", "_", texto)  # Reemplaza caracteres no alfanuméricos por "_"
    return texto.strip().lower()

# Leer los archivos
df_cierre = pd.read_excel("data/Forms Cierre de Voluntariado.xlsx", sheet_name="Sheet1")
df_bienvenida = pd.read_excel("data/Te damos la bienvenida__Dirección de Cultura Organizacional y Talento Humano.xlsx", sheet_name="data2025")

# Renombrar columnas con versión normalizada
df_cierre.columns = [normalizar(col) for col in df_cierre.columns]
df_bienvenida.columns = [normalizar(col) for col in df_bienvenida.columns]

# Nombres de columnas esperadas normalizadas
col_dni = "documento_de_identidad_dni_pasaporte_cedula"
col_fecha_cierre = "fecha_de_vinculacion_a_crea_peru"
col_fecha_bienvenida = "cual_es_tu_fecha_de_inicio_en_crea"

# Asegurar que las columnas estén como string
df_cierre[col_dni] = df_cierre[col_dni].astype(str).str.strip()
df_bienvenida[col_dni] = df_bienvenida[col_dni].astype(str).str.strip()

# Unir las dos tablas por el DNI
df_merged = df_cierre.merge(df_bienvenida[[col_dni, col_fecha_bienvenida]], on=col_dni, how="left")

# Función para comparar fechas
def comparar_fechas(row):
    fecha_cierre = str(row[col_fecha_cierre]).strip()
    fecha_bienvenida = str(row[col_fecha_bienvenida]).strip()
    if pd.isna(fecha_cierre) or pd.isna(fecha_bienvenida) or fecha_cierre == '' or fecha_bienvenida == '':
        return "INFORMACIÓN INCOMPLETA"
    elif fecha_cierre == fecha_bienvenida:
        return "COINCIDEN"
    else:
        return f"{fecha_cierre} ≠ {fecha_bienvenida}"

# Aplicar la comparación
df_merged["OBS_FECHA_INICIO"] = df_merged.apply(comparar_fechas, axis=1)

# Guardar archivo con resultados
fecha_hoy = datetime.today().strftime('%Y-%m-%d')
archivo_salida = f"TEST_salida/resultado_observaciones_{fecha_hoy}.xlsx"
df_merged.to_excel(archivo_salida, index=False)

print(f"✅ Archivo guardado: {archivo_salida}")

