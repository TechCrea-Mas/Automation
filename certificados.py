import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from datetime import datetime
import os
import glob
from dateutil.relativedelta import relativedelta  # útil para meses y días

# Buscar el archivo filtrado más reciente automáticamente
lista_archivos = glob.glob("TEST_salida/DNI_resultado_comparacion_filtrado_*.xlsx")
if not lista_archivos:
    raise FileNotFoundError("No se encontró ningún archivo filtrado en TEST_salida/")
ARCHIVO_FILTRADO = max(lista_archivos, key=os.path.getctime)
print(f"Usando archivo filtrado más reciente: {ARCHIVO_FILTRADO}")

CARPETA_CERTIFICADOS = "TEST_salida/certificados_pdf"
os.makedirs(CARPETA_CERTIFICADOS, exist_ok=True)

df = pd.read_excel(ARCHIVO_FILTRADO)
df_certificados = df[df["CERTIFICADO"] == "SI"]

# --- CAMPOS USADOS ---
campos = [
    "NOMBRE_SUNAT",
    "DNI",
    "Fecha de vinculación a Crea+ Perú:",
    "Fecha de desvinculación a Crea+ Perú:",
    "¿Qué rol desarrollaste dentro de la organización?",
]

# Función para calcular meses y días de voluntariado
def calcular_tiempo(inicio, fin):
    inicio = pd.to_datetime(inicio, dayfirst=True, errors="coerce")
    fin = pd.to_datetime(fin, dayfirst=True, errors="coerce")
    if pd.isna(inicio) or pd.isna(fin):
        return "un periodo no especificado"
    
    diff = relativedelta(fin, inicio)
    meses = diff.years * 12 + diff.months
    dias = diff.days
    
    if meses > 0 and dias > 0:
        return f"{meses} meses y {dias} días de voluntariado"
    elif meses > 0:
        return f"{meses} meses de voluntariado"
    elif dias > 0:
        return f"{dias} días de voluntariado"
    else:
        return "menos de 1 día de voluntariado"

# Fecha actual formateada
def formato_fecha_actual():
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]
    hoy = datetime.today()
    return f"Lima, {hoy.day} de {meses[hoy.month-1]} de {hoy.year}"

def generar_pdf(data, nombre_archivo):
    c = canvas.Canvas(nombre_archivo, pagesize=A4)
    width, height = A4
    margen = 40

    # Encabezado derecho
    c.setFont("Helvetica", 10)
    c.drawRightString(width-margen, height-margen, formato_fecha_actual())

    # Título central
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width/2, height-margen-40, "CERTIFICADO DE VOLUNTARIADO")

    # Calcular tiempo
    tiempo_voluntariado = calcular_tiempo(
        data["Fecha de vinculación a Crea+ Perú:"],
        data["Fecha de desvinculación a Crea+ Perú:"]
    )

    # Cuerpo del certificado
    c.setFont("Helvetica", 12)
    y = height-margen-80
    texto = (
        f'CREA MÁS PERU (en adelante, Crea+) es una asociación civil sin fines de lucro compuesta '
        'por un equipo multidisciplinario de jóvenes, el cual busca transformar la sociedad a través '
        'de una transformación personal de beneficiarios y voluntarios, otorgando herramientas para el crecimiento '
        'a través de un voluntariado profesional.\n\n'
        f'Mediante el presente, Crea+ deja constancia que "{data["NOMBRE_SUNAT"]}" con DNI "{data["DNI"]}", '
        f'participó como voluntaria/o desde el "{data["Fecha de vinculación a Crea+ Perú:"]}" al "{data["Fecha de desvinculación a Crea+ Perú:"]}" '
        f'en el rol de "{data["¿Qué rol desarrollaste dentro de la organización?"]}" cumpliendo con {tiempo_voluntariado}.\n\n'
        f'Certificamos que "{data["NOMBRE_SUNAT"]}" demostró responsabilidad y compromiso en el desarrollo de sus funciones.\n\n'
        'Se expide el presente certificado para los fines que se estimen convenientes.\n\n'
        'Atentamente,'
    )

    for linea in texto.split('\n'):
        c.drawString(margen, y, linea)
        y -= 18

    c.showPage()
    c.save()

# --- GENERACIÓN DE PDFS ---
for idx, row in df_certificados.iterrows():
    nombre = row["NOMBRE_SUNAT"].replace(" ", "_")
    nombre_pdf = f'{CARPETA_CERTIFICADOS}/certificado_{nombre}_{row["DNI"]}.pdf'
    generar_pdf(row, nombre_pdf)
    print(f"✅ Certificado generado: {nombre_pdf}")
