from reportlab.platypus import Paragraph, Frame, Spacer
from reportlab.pdfgen import canvas 
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import pandas as pd
import os, glob
from datetime import datetime
from dateutil.relativedelta import relativedelta

# ---------------------
# Funciones auxiliares
# ---------------------
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

def fecha_en_palabras(fecha_str):
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]
    fecha = pd.to_datetime(fecha_str, dayfirst=True, errors="coerce")
    if pd.isna(fecha):
        return "fecha no especificada"
    return f"{fecha.day} de {meses[fecha.month-1]} del {fecha.year}"

def formato_fecha_actual():
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]
    hoy = datetime.today()
    return f"Lima, {hoy.day} de {meses[hoy.month-1]} del {hoy.year}"

# ---------------------
# Generador de PDF
# ---------------------
def generar_pdf(data, nombre_archivo):
    w, h = A4
    c = canvas.Canvas(nombre_archivo, pagesize=A4)

    # Fondo con la plantilla institucional (JPG recomendado)
    plantilla_path = "plantilla_certificado.jpg"
    if os.path.exists(plantilla_path):
        fondo = ImageReader(plantilla_path)
        c.drawImage(fondo, 0, 0, width=w, height=h)

    # Logo institucional (ajusta posición y tamaño)
    logo_path = "assets/logo_crea.png"
    if os.path.exists(logo_path):
        logo = ImageReader(logo_path)
        c.drawImage(logo, w/2-75, h-110, width=150, height=50, mask='auto')

    # Fecha actual
    c.setFont("Helvetica", 12)
    c.drawString(w-250, h-140, formato_fecha_actual())

    # Título centrado
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(w/2, h-170, "CERTIFICADO DE VOLUNTARIADO")

    # Texto institucional
    texto_intro = (
        "<b>CREA MÁS PERÚ</b> (en adelante, Crea+) es una asociación civil sin fines de lucro compuesta "
        "por un equipo multidisciplinario de jóvenes, el cual busca transformar la sociedad a través de una transformación personal de beneficiarios y voluntarios, otorgando herramientas para el crecimiento a través de un voluntariado profesional."
    )

    texto_principal = (
        f"Mediante el presente, Crea+ deja constancia que <b>{data['NOMBRE_SUNAT']}</b> "
        f"con DNI <b>{data['DNI']}</b>, participó como voluntaria/o desde el <b>{fecha_en_palabras(data['Fecha de vinculación a Crea+ Perú:'])}</b> "
        f"al <b>{fecha_en_palabras(data['Fecha de desvinculación a Crea+ Perú:'])}</b> en el rol de <b>{data['¿Qué rol desarrollaste dentro de la organización?']}</b>, "
        f"cumpliendo con {calcular_tiempo(data['Fecha de vinculación a Crea+ Perú:'],data['Fecha de desvinculación a Crea+ Perú:'])}."
    )

    texto_final = (
        f"Certificamos que <b>{data['NOMBRE_SUNAT']}</b> demostró responsabilidad y compromiso en el desarrollo de sus funciones.<br/><br/>"
        "Se expide el presente certificado para los fines que se estimen convenientes.<br/><br/>"
        "Atentamente,"
    )

    # Estilos
    styles = getSampleStyleSheet()
    styleN = styles["Normal"]
    styleN.fontName = "Helvetica"
    styleN.fontSize = 12
    styleN.leading = 16

    # Marco para los párrafos (ajusta posición y tamaño según plantilla)
    frame_textos = Frame(70, h-350, w-140, 130, showBoundary=0)
    P_intro = Paragraph(texto_intro, styleN)
    P_principal = Paragraph(texto_principal, styleN)
    frame_textos.addFromList([P_intro, Spacer(1,12), P_principal], c)

    frame_final = Frame(70, h-440, w-140, 80, showBoundary=0)
    P_final = Paragraph(texto_final, styleN)
    frame_final.addFromList([P_final], c)

    # Firma + sello (ajusta posiciones y tamaños según plantilla)
    firma_path = "assets/firma.png"
    sello_path = "assets/pie_pagina.png"
    if os.path.exists(firma_path):
        c.drawImage(firma_path, w/2-120, 125, width=90, height=45, mask='auto')
    if os.path.exists(sello_path):
        c.drawImage(sello_path, w/2+30, 125, width=90, height=45, mask='auto')

    # Nombre y cargo
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(w/2, 110, "Diego Cabrera Zanatta")
    c.setFont("Helvetica", 10)
    c.drawCentredString(w/2, 97, "Coordinador de Gestión de Talento Humano")

    c.save()

# ---------------------
# Bucle principal
# ---------------------
lista_archivos = glob.glob("TEST_salida/DNI_resultado_comparacion_filtrado_*.xlsx")
if not lista_archivos:
    raise FileNotFoundError("No se encontró ningún archivo filtrado en TEST_salida/")
ARCHIVO_FILTRADO = max(lista_archivos, key=os.path.getctime)

df = pd.read_excel(ARCHIVO_FILTRADO)
df_certificados = df[df["CERTIFICADO"] == "SI"]

CARPETA_CERTIFICADOS = "TEST_salida/certificados_pdf"
os.makedirs(CARPETA_CERTIFICADOS, exist_ok=True)

for _, row in df_certificados.iterrows():
    nombre = row["NOMBRE_SUNAT"].replace(" ", "_")
    nombre_pdf = f'{CARPETA_CERTIFICADOS}/certificado_{nombre}_{row["DNI"]}.pdf'
    generar_pdf(row, nombre_pdf)
