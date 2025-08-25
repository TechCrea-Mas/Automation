#!/usr/bin/env python
# coding: utf-8

import pandas as pd
from pathlib import Path
from datetime import datetime
import os

# Crear carpeta de salida si no existe
Path("TEST_salida").mkdir(exist_ok=True)

# Leer las hojas desde los archivos
df_cierre = pd.read_excel("data/Forms Cierre de Voluntariado.xlsx", sheet_name="Sheet1")
df_bienvenida = pd.read_excel("data/Te damos la bienvenida__Dirección de Cultura Organizacional y Talento Humano.xlsx", sheet_name="data2025")
df_DNI_SUNAT = pd.read_excel("data/DNI_resultado_observaciones.xlsx", sheet_name="Sheet1")

col_dni_original = "Documento de identidad (DNI/Pasaporte/Cédula):\n"
col_fecha_cierre = "Fecha de vinculación a Crea+ Perú:\n"
col_fecha_bienvenida = "¿Cuál es tu fecha de inicio en Crea+?"

# ✏️ Renombrar columnas para simplificar
df_cierre = df_cierre.rename(columns={col_dni_original: "DNI"})
df_bienvenida = df_bienvenida.rename(columns={col_dni_original.strip(): "DNI"})

# 🔍 Normalizar valores
df_cierre["DNI"] = df_cierre["DNI"].astype(str).str.strip()
df_bienvenida["DNI"] = df_bienvenida["DNI"].astype(str).str.strip()

# 🔗 Unir las dos tablas por el DNI
df_merged = df_cierre.merge(
    df_bienvenida[["DNI", col_fecha_bienvenida]],
    on="DNI",
    how="left"
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

# ==============================
# --- INICIO BLOQUE SOLICITADO ---
# ==============================

# Normalizar columna DNI para unir con la consulta SUNAT
df_merged["DNI"] = df_merged["DNI"].astype(str).str.zfill(8)
df_DNI_SUNAT["DNI"] = df_DNI_SUNAT["DNI"].astype(str).str.zfill(8)

# Unimos df_merged con df_DNI_SUNAT para traer datos de SUNAT
df_final = df_merged.merge(
    df_DNI_SUNAT,
    left_on="DNI",
    right_on="DNI",
    how="left"
)

# 1. Filtrar solo los casos OK & COINCIDEN
df_filtrado = df_final[
    (df_final["ESTADO_SUNAT"] == "✅ OK") &
    (df_final["OBS_FECHA_INICIO"] == "COINCIDEN")
].copy()

# 2. Concatenar nombres y apellidos del archivo original
# Cambia estos nombres de columna si tus columnas son diferentes
col_nombres = "Nombres completos"
col_apellidos = "Apellidos completos"

if col_nombres in df_filtrado.columns and col_apellidos in df_filtrado.columns:
    df_filtrado["NOMBRE_COMPLETO_EXCEL"] = (
        df_filtrado[col_nombres].astype(str).str.strip() + " " +
        df_filtrado[col_apellidos].astype(str).str.strip()
    ).str.upper().str.replace(r"\s+", " ", regex=True)
else:
    # Si las columnas no existen, crea columna vacía para evitar error
    df_filtrado["NOMBRE_COMPLETO_EXCEL"] = ""

# 3. Normalizar NOMBRE_SUNAT y comparar
df_filtrado["NOMBRE_SUNAT_NORMALIZADO"] = (
    df_filtrado["NOMBRE_SUNAT"].astype(str).str.upper().str.strip().str.replace(r"\s+", " ", regex=True)
)

def comparar_nombres(row):
    if pd.isna(row["NOMBRE_COMPLETO_EXCEL"]) or pd.isna(row["NOMBRE_SUNAT_NORMALIZADO"]):
        return "INCOMPLETO"
    return "COINCIDEN" if row["NOMBRE_COMPLETO_EXCEL"] == row["NOMBRE_SUNAT_NORMALIZADO"] else "NO COINCIDEN"

df_filtrado["OBS_NOMBRE_SUNAT"] = df_filtrado.apply(comparar_nombres, axis=1)

# 4. Guardar archivo filtrado y comparado
fecha = datetime.now().strftime("%Y-%m-%d_%H-%M")
DEBUG_FOLDER = "TEST_salida"
nombre_archivo_final = os.path.join(DEBUG_FOLDER, f"DNI_resultado_filtrado_{fecha}.xlsx")
df_filtrado.to_excel(nombre_archivo_final, index=False)
print(f"✅ Archivo filtrado y comparado por nombre guardado en: {nombre_archivo_final}")

# ==============================
# --- FIN BLOQUE SOLICITADO ---
# ==============================

# Guardar archivo con resultados generales (por si también quieres el archivo general)
fecha_hoy = datetime.today().strftime('%Y-%m-%d')
archivo_salida = f"TEST_salida/resultado_observaciones_{fecha_hoy}.xlsx"
df_merged.to_excel(archivo_salida, index=False)

print(f"✅ Archivo guardado: {archivo_salida}")
