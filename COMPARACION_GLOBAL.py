import pandas as pd
from pathlib import Path
from datetime import datetime

# Carpeta y patrón de archivo
DIR_SALIDA = "TEST_salida"
# Busca el archivo más reciente con el patrón deseado
from glob import glob
import os

# Encuentra el archivo más reciente con el patrón
archivos = sorted(
    glob(os.path.join(DIR_SALIDA, "DNI_resultado_comparacion_*.xlsx")),
    key=os.path.getmtime,
    reverse=True
)
if not archivos:
    raise FileNotFoundError("No se encontró ningún archivo con patrón 'DNI_resultado_comparacion_*.xlsx' en TEST_salida")
archivo = archivos[0]

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

# Cargar archivo y seleccionar solo las columnas indicadas (las que existan)
df = pd.read_excel(archivo)
columnas_existentes = [col for col in columnas_requeridas if col in df.columns]
df_filtrado = df[columnas_existentes]

# Guardar el archivo filtrado
fecha_archivo = datetime.now().strftime("%Y-%m-%d_%H-%M")
archivo_filtrado = f"{DIR_SALIDA}/DNI_resultado_comparacion_filtrado_{fecha_archivo}.xlsx"
df_filtrado.to_excel(archivo_filtrado, index=False)
print(f"✅ Descarga lista: {archivo_filtrado}")
