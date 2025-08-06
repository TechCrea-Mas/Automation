#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#!/usr/bin/env python
# coding: utf-8

# In[ ]:

import pandas as pd
from pathlib import Path
from datetime import datetime 


# Crear carpeta de salida si no existe
Path("TEST_salida").mkdir(exist_ok=True)

# Leer las hojas desde los archivos
df_cierre = pd.read_excel("data/Forms Cierre de Voluntariado.xlsx", sheet_name="Sheet1")
df_bienvenida = pd.read_excel("data/Te damos la bienvenida__Dirección de Cultura Organizacional y Talento Humano.xlsx", sheet_name="data2025")

col_dni = "Documento de identidad (DNI/Pasaporte/Cédula):\n"
col_fecha_cierre = "Fecha de vinculación a Crea+ Perú:\n"
col_fecha_bienvenida = "¿Cuál es tu fecha de inicio en Crea+?"

# 🔍 Normalizar valores
df_cierre[col_dni] = df_cierre[col_dni].astype(str).str.strip()
df_bienvenida[col_dni.strip()] = df_bienvenida[col_dni.strip()].astype(str).str.strip()

# Unir las dos tablas por el DNI
df_merged = df_cierre.merge(df_bienvenida[[col_dni.strip(), col_fecha_bienvenida]], left_on=col_dni, right_on=col_dni.strip(), how="left"
)

# 🧠 Comparar fechas
def comparar_fechas(row):
    fecha_cierre = str(row[col_fecha_cierre]).strip()
    fecha_bienvenida = str(row[col_fecha_bienvenida]).strip()
    if pd.isna(fecha_cierre) or pd.isna(fecha_bienvenida) or fecha_cierre == '' or fecha_bienvenida == '':
        return "INFORMACIÓN INCOMPLETA"
    elif fecha_cierre == fecha_bienvenida:
        return "COINCIDEN"
    else:
        return f"{fecha_cierre} ≠ {fecha_bienvenida}"

# ➕ Crear columna de observaciones
df_merged["OBS_FECHA_INICIO"] = df_merged.apply(comparar_fechas, axis=1)

# Guardar archivo con resultados
fecha_hoy = datetime.today().strftime('%Y-%m-%d')
archivo_salida = f"TEST_salida/resultado_observaciones_{fecha_hoy}.xlsx"
df_merged.to_excel(archivo_salida, index=False)

print(f"✅ Archivo guardado: {archivo_salida}")

