import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from datetime import datetime
import os

# --- CONFIGURACIONES ---
ARCHIVO_FILTRADO = "TEST_salida/DNI_resultado_comparacion_filtrado_YYYY-MM-DD_HH-MM.xlsx"  # cambia por el nombre real generado
CARPETA_CERTIFICADOS = "TEST_salida/certificados_pdf"
os.makedirs(CARPETA_CERTIFICADOS, exist_ok=True)

# --- CARGA Y FILTRO DE DATOS ---
df = pd.read_excel(ARCHIVO_FILTRADO)
df_certificados = df[df["CERTIFICADO"] == "SI"]

# --- CAMPOS USADOS ---
campos = [
    "NOMBRE_SUNAT",
    "DNI",
    "Fecha de vinculación a Crea+ Perú:",
    "¿Qué rol desarrollaste dentro de la organización?",
    # Si tienes campos de horas/meses, agrégalos aquí:
    # "X horas / X meses"
]

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

    # Cuerpo del certificado
    c.setFont("Helvetica", 12)
    y = height-margen-80
    texto = (f'CREA MÁS PERU (en adelante, Crea+) es una asociación civil sin fines de lucro compuesta '
             'por un equipo multidisciplinario de jóvenes, el cual busca transformar la sociedad a través '
             'de una transformación personal de beneficiarios y voluntarios, otorgando herramientas para el crecimiento '
             'a través de un voluntariado profesional.\n\n'
             f'Mediante el presente, Crea+ deja constancia que "{data["NOMBRE_SUNAT"]}" con DNI "{data["DNI"]}", '
             f'participó como parte del equipo desde el "{data["Fecha de vinculación a Crea+ Perú:"]}" en el rol de '
             f'"{data["¿Qué rol desarrollaste dentro de la organización?"]}". '
             # Puedes agregar aquí cálculo de horas/meses si tienes los datos
             '\n\nCertificamos que "{data["NOMBRE_SUNAT"]}" demostró responsabilidad y compromiso en el desarrollo de sus funciones.\n\n'
             'Se expide el presente certificado para los fines que se estimen convenientes.\n\n'
             'Atentamente,'
    )

    for linea in texto.split('\n'):
        c.drawString(margen, y, linea)
        y -= 18

    # Firma opcional: puedes añadir imagen de firma/logo si lo deseas
    # firma = ImageReader("firma.png")
    # c.drawImage(firma, margen, y-50, width=150, height=50, mask='auto')

    c.showPage()
    c.save()

# --- GENERACIÓN DE PDFS ---
for idx, row in df_certificados.iterrows():
    nombre = row["NOMBRE_SUNAT"].replace(" ", "_")
    nombre_pdf = f'{CARPETA_CERTIFICADOS}/certificado_{nombre}_{row["DNI"]}.pdf'
    generar_pdf(row, nombre_pdf)
    print(f"✅ Certificado generado: {nombre_pdf}")
