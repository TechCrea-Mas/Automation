#!/usr/bin/env python
# coding: utf-8

import pandas as pd
from pathlib import Path
from datetime import datetime
import os
import sys
import unicodedata

# Crear carpeta de salida si no existe
Path("TEST_salida").mkdir(exist_ok=True)

# === ARCHIVOS DE ENTRADA ===
fecha_hoy = datetime.today().strftime('%Y-%m-%d')
archivo_cierre = "data/Forms Cierre de Voluntariado.xlsx"
archivo_bienvenida = "data/Te damos la bienvenida__Dirección de Cultura Organizacional y Talento Humano.xlsx"
archivo_sunat = "data/DNI_OBS.xlsx"

# === LECTURA ===
df_cierre = pd.read_excel(archivo_cierre, sheet_name="Sheet1")
df_bienvenida = pd.read_excel(archivo_bienvenida, sheet_name="data2025")
df_sunat = pd.read_excel(archivo_sunat, sheet_name="Sheet1")

# --- Normalizar nombres de columnas ---
df_cierre.columns = df_cierre.columns.str.strip().str.replace("\n", "")
df_bienvenida.columns = df_bienvenida.columns.str.strip().str.replace("\n", "")
df_sunat.columns = df_sunat.columns.str.strip().str.replace("\n", "")

# --- Columnas clave ---
col_dni = "Documento de identidad (DNI/Pasaporte/Cédula):"
col_fecha_cierre = "Fecha de vinculación a Crea+ Perú:"
col_fecha_bienvenida = "¿Cuál es tu fecha de inicio en Crea+?"

# Renombrar para simplificar
df_cierre = df_cierre.rename(columns={col_dni: "DNI"})
df_bienvenida = df_bienvenida.rename(columns={col_dni: "DNI"})
df_sunat = df_sunat.rename(columns={"DNI": "DNI"})

# Normalizar DNI
for df in [df_cierre, df_bienvenida, df_sunat]:
    df["DNI"] = df["DNI"].astype(str).str.strip().str.zfill(8)

# --- Merge cierre + bienvenida ---
df_merge = df_cierre.merge(
    df_bienvenida[["DNI", col_fecha_bienvenida]],
    on="DNI",
    how="left"
)

# --- Comparar fechas ---
def comparar_fechas(row):
    f1 = str(row[col_fecha_cierre]).strip()
    f2 = str(row[col_fecha_bienvenida]).strip()
    if pd.isna(f1) or pd.isna(f2) or f1 == "" or f2 == "":
        return "INFORMACIÓN INCOMPLETA"
    return "COINCIDEN" if f1 == f2 else f"{f1} ≠ {f2}"

df_merge["OBS_FECHA_INICIO"] = df_merge.apply(comparar_fechas, axis=1)

# --- Merge con SUNAT ---
df_final = df_merge.merge(
    df_sunat[["DNI", "NOMBRE_SUNAT", "ESTADO_SUNAT"]],
    on="DNI",
    how="left"
)

# --- Construir nombre completo desde CIERRE ---
df_final["NOMBRE_COMPLETO_EXCEL"] = (
    df_final["Nombres completos"].astype(str).str.strip() + " " +
    df_final["Apellidos completos"].astype(str).str.strip()
).str.upper().str.replace(r"\s+", " ", regex=True)

df_final["NOMBRE_SUNAT_NORMALIZADO"] = (
    df_final["NOMBRE_SUNAT"].astype(str).str.upper().str.strip().str.replace(r"\s+", " ", regex=True)
)

# --- Comparar nombres ignorando orden ---
def nombres_coinciden(nombre1, nombre2):
    if not nombre1 or not nombre2:
        return "INCOMPLETO"
    def normalizar(txt):
        txt = unicodedata.normalize("NFD", txt)
        txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
        return txt.upper().strip()
    set1, set2 = set(normalizar(nombre1).split()), set(normalizar(nombre2).split())
    return "COINCIDEN" if set1 == set2 else "NO COINCIDEN"

df_final["OBS_NOMBRE_SUNAT"] = df_final.apply(
    lambda row: nombres_coinciden(row["NOMBRE_COMPLETO_EXCEL"], row["NOMBRE_SUNAT_NORMALIZADO"]),
    axis=1
)

# --- Seleccionar columnas en el orden final ---
columnas_finales = [
    "Id","Hora de inicio","Hora de finalización","Correo electrónico","Nombre",
    "Nombres completos","Apellidos completos","DNI","Celular de contacto:","Correo electrónico:",
    "¿Qué rol desarrollaste dentro de la organización?","Fecha de desvinculación a Crea+ Perú:",
    "Fecha de vinculación a Crea+ Perú:","¿Cuál fue el motivo de tu salida?","Capacitación inicial",
    "Acompañamiento y apoyo  de los líderes durante el voluntariado","Claridad en las tareas asignadas",
    "Recursos y herramientas disponibles","Ambiente de trabajo","Motivación recibida en la Asamblea de Impacto",
    "Puntualidad en la asistencia a actividades y reuniones","Satisfacción general con la experiencia",
    "¿Qué aprendiste durante tu voluntariado?","¿Qué mejorarías para futuros voluntarios?",
    "¿Recomendarías este programa de voluntariado a otras personas?","¿Te gustaría seguir vinculado a la organización?",
    "Protección de datos","REGISTRO DE ENTREGA",col_fecha_bienvenida,"OBS_FECHA_INICIO",
    "OBS_NOMBRE_SUNAT","NOMBRE_SUNAT","ESTADO_SUNAT"
]

df_salida = df_final[columnas_finales].copy()

# --- Guardar resultado único ---
fecha_archivo = datetime.now().strftime("%Y-%m-%d_%H-%M")
archivo_resultado = f"TEST_salida/TEST_resultado_comparacion_{fecha_archivo}.xlsx"
df_salida.to_excel(archivo_resultado, index=False)

print(f"✅ Archivo final generado: {archivo_resultado}")
