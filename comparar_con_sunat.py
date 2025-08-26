#!/usr/bin/env python
# coding: utf-8

import pandas as pd
from pathlib import Path
from datetime import datetime
import os
import sys

# Crear carpeta de salida si no existe
Path("TEST_salida").mkdir(exist_ok=True)

# === CONFIGURACIÓN DE ARCHIVOS ===
# Archivo generado previamente por el script anterior
fecha_hoy = datetime.today().strftime('%Y-%m-%d')
archivo_base = f"TEST_salida/resultado_observaciones_{fecha_hoy}.xlsx"
archivo_sunat = pd.read_excel("data/DNI_OBS.xlsx", sheet_name="Sheet1"

# === VERIFICACIÓN DE EXISTENCIA ===
if not (os.path.exists(archivo_base) and os.path.exists(archivo_sunat)):
    print(f"❌ Faltan archivos de entrada:")
    print(f" - {archivo_base}: {'✅' if os.path.exists(archivo_base) else '❌ NO ENCONTRADO'}")
    print(f" - {archivo_sunat}: {'✅' if os.path.exists(archivo_sunat) else '❌ NO ENCONTRADO'}")
    print("Ejecuta primero el script generador y asegúrate de que ambos archivos existan.")
    sys.exit(1)

# === LECTURA DE ARCHIVOS ===
df_base = pd.read_excel(archivo_base, sheet_name="Sheet1")
df_sunat = pd.read_excel(archivo_sunat, sheet_name="Sheet1")

# === NORMALIZACIÓN DE DNI ===
df_base["DNI"] = df_base["DNI"].astype(str).str.zfill(8)
df_sunat["DNI"] = df_sunat["DNI"].astype(str).str.zfill(8)

# === UNIÓN DE DATAFRAMES POR DNI ===
df_final = df_base.merge(
    df_sunat,
    on="DNI",
    how="left",
    suffixes=('', '_SUNAT')
)

# === FILTRO: Solo los que tienen ESTADO_SUNAT OK y OBS_FECHA_INICIO COINCIDEN ===
df_filtrado = df_final[
    (df_final["ESTADO_SUNAT"] == "✅ OK") &
    (df_final["OBS_FECHA_INICIO"] == "COINCIDEN")
].copy()

# === CONCATENACIÓN Y NORMALIZACIÓN DE NOMBRES ===
col_nombres = "Nombres completos"
col_apellidos = "Apellidos completos"
if col_nombres in df_filtrado.columns and col_apellidos in df_filtrado.columns:
    df_filtrado["NOMBRE_COMPLETO_EXCEL"] = (
        df_filtrado[col_nombres].astype(str).str.strip() + " " +
        df_filtrado[col_apellidos].astype(str).str.strip()
    ).str.upper().str.replace(r"\s+", " ", regex=True)
else:
    df_filtrado["NOMBRE_COMPLETO_EXCEL"] = ""

df_filtrado["NOMBRE_SUNAT_NORMALIZADO"] = (
    df_filtrado["NOMBRE_SUNAT"].astype(str).str.upper().str.strip().str.replace(r"\s+", " ", regex=True)
)

def comparar_nombres(row):
    if pd.isna(row["NOMBRE_COMPLETO_EXCEL"]) or pd.isna(row["NOMBRE_SUNAT_NORMALIZADO"]):
        return "INCOMPLETO"
    return "COINCIDEN" if row["NOMBRE_COMPLETO_EXCEL"] == row["NOMBRE_SUNAT_NORMALIZADO"] else "NO COINCIDEN"

df_filtrado["OBS_NOMBRE_SUNAT"] = df_filtrado.apply(comparar_nombres, axis=1)

# === GUARDAR RESULTADO FINAL ===
fecha_archivo = datetime.now().strftime("%Y-%m-%d_%H-%M")
archivo_resultado = f"TEST_salida/DNI_resultado_comparacion_{fecha_archivo}.xlsx"
df_filtrado.to_excel(archivo_resultado, index=False)
print(f"✅ Archivo final comparado guardado en: {archivo_resultado}")
