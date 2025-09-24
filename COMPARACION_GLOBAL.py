#!/usr/bin/env python
# coding: utf-8

import pandas as pd
from pathlib import Path
from datetime import datetime
import os
import sys
import unicodedata
from glob import glob

# === CONFIGURACIÓN DE RUTAS Y VARIABLES ===
DIR_SALIDA = "TEST_salida"
Path(DIR_SALIDA).mkdir(exist_ok=True)

# Archivos de entrada
archivo_cierre = "data/Cierre de rol.xlsx"
archivo_bienvenida = "data/Forms Bienvenida al Creando.xlsx"
archivo_sunat = "data/2DNI_OBS.xlsx"

# Hojas y columnas relevantes
hoja_cierre = "salida2025"
hoja_bienvenida = "Bienvenida"
hoja_sunat = "Sheet1"

col_dni_original = "Documento de identidad (DNI/Pasaporte/Cédula):\n"
col_fecha_cierre = "Fecha de vinculación a Crea+ Perú:\n"
col_fecha_bienvenida = "¿Cuál es tu fecha de inicio en Crea+?"
col_nombres = "Nombres completos"
col_apellidos = "Apellidos completos"

# === FUNCIONES DE NORMALIZACIÓN DE NOMBRES ===
def normalizar_nombre(nombre):
    if pd.isna(nombre):
        return ''
    nombre = str(nombre).upper()
    nombre = ''.join(
        c for c in unicodedata.normalize('NFD', nombre)
        if unicodedata.category(c) != 'Mn'
    )
    nombre = ''.join(c for c in nombre if c.isalpha() or c.isspace())
    nombre = ' '.join(nombre.split())
    palabras = nombre.split()
    palabras.sort()
    return ' '.join(palabras)

# === LECTURA Y PROCESAMIENTO DE ARCHIVOS DE VOLUNTARIADO Y BIENVENIDA ===
try:
    df_cierre = pd.read_excel(archivo_cierre, sheet_name=hoja_cierre)
    df_bienvenida = pd.read_excel(archivo_bienvenida, sheet_name=hoja_bienvenida)
except Exception as e:
    print(f"❌ Error leyendo archivos de voluntariado/bienvenida: {e}")
    sys.exit(1)

# Renombrar y normalizar columnas DNI
df_cierre = df_cierre.rename(columns={col_dni_original: "DNI"})
df_bienvenida = df_bienvenida.rename(columns={col_dni_original.strip(): "DNI"})

df_cierre["DNI"] = df_cierre["DNI"].astype(str).str.strip().str.zfill(8)
df_bienvenida["DNI"] = df_bienvenida["DNI"].astype(str).str.strip().str.zfill(8)

# Merge y comparación de fechas
df_merged = df_cierre.merge(
    df_bienvenida[["DNI", col_fecha_bienvenida]],
    on="DNI",
    how="left"
)

def comparar_fechas(row):
    fecha_cierre = str(row.get(col_fecha_cierre, "")).strip()
    fecha_bienvenida = str(row.get(col_fecha_bienvenida, "")).strip()
    if pd.isna(fecha_cierre) or pd.isna(fecha_bienvenida) or fecha_cierre == '' or fecha_bienvenida == '':
        return "INFORMACIÓN INCOMPLETA"
    elif fecha_cierre == fecha_bienvenida:
        return "COINCIDEN"
    else:
        return f"{fecha_cierre} ≠ {fecha_bienvenida}"

df_merged["OBS_FECHA_INICIO"] = df_merged.apply(comparar_fechas, axis=1)

# === LECTURA Y CRUCE CON SUNAT ===
fecha_hoy = datetime.today().strftime('%Y-%m-%d')
archivo_temp = f"{DIR_SALIDA}/resultado_observaciones_{fecha_hoy}.xlsx"
df_merged.to_excel(archivo_temp, index=False)

if not os.path.exists(archivo_temp) or not os.path.exists(archivo_sunat):
    print("❌ Faltan archivos de entrada para cruce con SUNAT:")
    print(f" - {archivo_temp}: {'✅' if os.path.exists(archivo_temp) else '❌ NO ENCONTRADO'}")
    print(f" - {archivo_sunat}: {'✅' if os.path.exists(archivo_sunat) else '❌ NO ENCONTRADO'}")
    sys.exit(1)

df_base = pd.read_excel(archivo_temp)
df_sunat = pd.read_excel(archivo_sunat, sheet_name=hoja_sunat)

# Normalizar nombres de columnas y DNI
df_base.columns = df_base.columns.str.strip().str.replace("\n", "")
df_sunat.columns = df_sunat.columns.str.strip().str.replace("\n", "")

df_base["DNI"] = df_base["DNI"].astype(str).str.zfill(8)
df_sunat["DNI"] = df_sunat["DNI"].astype(str).str.zfill(8)

cols_sunat = [c for c in df_sunat.columns if c not in df_base.columns or c == "DNI"]
df_final = df_base.merge(
    df_sunat[cols_sunat],
    on="DNI",
    how="left"
)

# === CONCATENACIÓN Y NORMALIZACIÓN DE NOMBRES (EN TODOS LOS REGISTROS) ===
nuevas_columnas = []

# NOMBRE_COMPLETO_EXCEL
if col_nombres in df_final.columns and col_apellidos in df_final.columns:
    nombre_completo_excel = (
        df_final[col_nombres].astype(str).str.strip() + " " +
        df_final[col_apellidos].astype(str).str.strip()
    ).str.upper().str.replace(r"\s+", " ", regex=True)
    df_final["NOMBRE_COMPLETO_EXCEL"] = nombre_completo_excel
else:
    df_final["NOMBRE_COMPLETO_EXCEL"] = ""
nuevas_columnas.append("NOMBRE_COMPLETO_EXCEL")

# NOMBRE_SUNAT_NORMALIZADO
if "NOMBRE_SUNAT" in df_final.columns:
    nombre_sunat_norm = (
        df_final["NOMBRE_SUNAT"].astype(str).str.upper().str.strip().str.replace(r"\s+", " ", regex=True)
    )
    df_final["NOMBRE_SUNAT_NORMALIZADO"] = nombre_sunat_norm
else:
    df_final["NOMBRE_SUNAT_NORMALIZADO"] = ""
nuevas_columnas.append("NOMBRE_SUNAT_NORMALIZADO")

# NOMBRE_COMPLETO_EXCEL_NORMALIZADO
df_final["NOMBRE_COMPLETO_EXCEL_NORMALIZADO"] = df_final["NOMBRE_COMPLETO_EXCEL"].apply(normalizar_nombre)
nuevas_columnas.append("NOMBRE_COMPLETO_EXCEL_NORMALIZADO")

# NOMBRE_SUNAT_ORDENADO
df_final["NOMBRE_SUNAT_ORDENADO"] = df_final["NOMBRE_SUNAT_NORMALIZADO"].apply(normalizar_nombre)
nuevas_columnas.append("NOMBRE_SUNAT_ORDENADO")

# OBS_NOMBRE_SUNAT
def comparar_nombres(row):
    nombre1 = row["NOMBRE_COMPLETO_EXCEL_NORMALIZADO"]
    nombre2 = row["NOMBRE_SUNAT_ORDENADO"]
    if not nombre1 or not nombre2:
        return "INCOMPLETO"
    return "COINCIDEN" if nombre1 == nombre2 else "NO COINCIDEN"

df_final["OBS_NOMBRE_SUNAT"] = df_final.apply(comparar_nombres, axis=1)
nuevas_columnas.append("OBS_NOMBRE_SUNAT")

# === COLUMNA CERTIFICADO SEGÚN CONDICIÓN DE COINCIDEN EN AMBAS OBSERVACIONES ===
def certificado_condicion(row):
    if row.get("OBS_FECHA_INICIO") == "COINCIDEN" and row.get("OBS_NOMBRE_SUNAT") == "COINCIDEN":
        return "SI"
    else:
        return "NO"

df_final["CERTIFICADO"] = df_final.apply(certificado_condicion, axis=1)
nuevas_columnas.append("CERTIFICADO")

# === REORDENAR COLUMNAS: NUEVAS COLUMNAS AL FINAL EN ORDEN ===
otras = [c for c in df_final.columns if c not in nuevas_columnas]
df_final = df_final[otras + nuevas_columnas]

# === GUARDAR RESULTADO FINAL (TODOS LOS REGISTROS) ===
fecha_archivo = datetime.now().strftime("%Y-%m-%d_%H-%M")
archivo_resultado = f"{DIR_SALIDA}/DNI_resultado_comparacion_{fecha_archivo}.xlsx"
df_final.to_excel(archivo_resultado, index=False)
print(f"✅ Archivo final comparado guardado en: {archivo_resultado}")

# === FILTRAR SOLO LAS COLUMNAS SOLICITADAS Y GUARDAR OTRA DESCARGA ===
# Las columnas requeridas
columnas_requeridas = [
    "Id",
    "Hora de inicio",
    "Hora de finalización",
    "Correo electrónico",
    "Nombre",
    "Nombres completos",
    "Apellidos completos",
    "DNI",
    "Celular de contacto:",
    "Correo electrónico:",
    "¿Qué rol desarrollaste dentro de la organización?",
    "¿En qué área o equipo participaste?"
    "Fecha de desvinculación a Crea+ Perú:",
    "Fecha de vinculación a Crea+ Perú:",
    "¿Cuál fue el motivo de tu salida?",
    "Capacitación inicial",
    "Acompañamiento y apoyo  de los líderes durante el voluntariado",
    "Claridad en las tareas asignadas",
    "Recursos y herramientas disponibles",
    "Ambiente de trabajo",
    "Motivación recibida en la Asamblea de Impacto",
    "Puntualidad en la asistencia a actividades y reuniones",
    "Satisfacción general con la experiencia",
    "¿Qué aprendiste durante tu voluntariado?",
    "¿Qué mejorarías para futuros voluntarios?",
    "¿Recomendarías este programa de voluntariado a otras personas?",
    "¿Te gustaría seguir vinculado a la organización?",
    "Protección de datos",
    "REGISTRO DE ENTREGA",
    "¿Cuál es tu fecha de inicio en Crea+?",
    "OBS_FECHA_INICIO",
    "NOMBRE_SUNAT",
    "OBS_NOMBRE_SUNAT",
    "ESTADO_SUNAT",
    "CERTIFICADO"
]

# Selecciona solo las columnas que existan en el resultado
columnas_existentes = [col for col in columnas_requeridas if col in df_final.columns]
df_filtrado = df_final[columnas_existentes]

archivo_filtrado = f"{DIR_SALIDA}/DNI_resultado_comparacion_filtrado_{fecha_archivo}.xlsx"
df_filtrado.to_excel(archivo_filtrado, index=False)
print(f"✅ Descarga lista: {archivo_filtrado}")
