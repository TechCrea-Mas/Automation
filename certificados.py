from reportlab.pdfgen import canvas 
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os, glob
import re

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

def formatear_nombre(texto):
    if pd.isna(texto):
        return ""
    return str(texto).title()

def generar_pdf(data, nombre_archivo):
    w, h = A4
    c = canvas.Canvas(nombre_archivo, pagesize=A4)
    
    # Fondo institucional
    plantilla_path = "plantilla_certificado.jpg"
    if os.path.exists(plantilla_path):
        fondo = ImageReader(plantilla_path)
        c.drawImage(fondo, 0, 0, width=w, height=h)

    # Estilos para ReportLab
    styles = getSampleStyleSheet()
    styleTitulo = ParagraphStyle(
        'titulo',
        parent=styles['Heading1'],
        alignment=1,  # centrado
        fontSize=16,
        fontName="Helvetica-Bold",
        spaceAfter=18
    )
    styleNormal = ParagraphStyle(
        'normal',
        fontName="Helvetica",
        fontSize=12,
        leading=16,
        alignment=4,  # justificado
        spaceAfter=12
    )

    # Fecha actual centrada arriba
    c.setFont("Helvetica", 12)
    c.drawCentredString(w/2, h-115, formato_fecha_actual())

    # Título centrado
    P_titulo = Paragraph("CERTIFICADO DE VOLUNTARIADO", styleTitulo)
    frame_titulo = Frame(0, h-160, w, 30, showBoundary=0)
    frame_titulo.addFromList([P_titulo], c)

    # Texto institucional y dinámico
    fecha_vinculacion = fecha_en_palabras(data["Fecha de vinculación a Crea+ Perú:"])
    fecha_desvinculacion = fecha_en_palabras(data["Fecha de desvinculación a Crea+ Perú:"])
    tiempo_voluntariado = calcular_tiempo(
        data["Fecha de vinculación a Crea+ Perú:"],
        data["Fecha de desvinculación a Crea+ Perú:"]
    )

    nombre = formatear_nombre(data["NOMBRE_SUNAT"])
    dni = str(int(data["DNI"])) if not pd.isna(data["DNI"]) else ""
    area = formatear_nombre(data["¿En qué área o equipo participaste?"])
    rol = formatear_nombre(data["¿Qué rol desarrollaste dentro de la organización?"])

    texto = (
        "<b>CREA MÁS PERU</b> (en adelante, Crea+) es una asociación civil sin fines de lucro compuesta por un equipo multidisciplinario de jóvenes, el cual busca transformar la sociedad a través de una transformación personal de beneficiarios y voluntarios, otorgando herramientas para el crecimiento a través de un voluntariado profesional.<br/><br/>"
        f"Mediante el presente, Crea+ deja constancia que <b>{nombre}</b> con DNI <b>{dni}</b>, participó como voluntaria/o en el área/equipo <b>{area}</b> desde el <b>{fecha_vinculacion}</b> al <b>{fecha_desvinculacion}</b> en el rol de <b>{rol}</b>, cumpliendo con <b>{tiempo_voluntariado}</b>.<br/><br/>"
        f"Certificamos que <b>{nombre}</b> demostró responsabilidad y compromiso en el desarrollo de sus funciones.<br/><br/>"
        "Se expide el presente certificado para los fines que se estimen convenientes.<br/><br/>"
        "Atentamente,"
    )

    # Frame principal (área blanca central, no choca con firma/sello)
    frame = Frame(80, 220, w-160, 330, showBoundary=0)  # left, bottom, width, height
    P = Paragraph(texto, styleNormal)
    frame.addFromList([P], c)

    # Firma, sello y pie de página YA están en la plantilla, no se agregan por código

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
    nombre = formatear_nombre(row["NOMBRE_SUNAT"]).replace(" ", "_")
    nombre = re.sub(r'[^\w\-]', '', nombre)  # elimina caracteres no válidos
    dni = str(int(row["DNI"])) if not pd.isna(row["DNI"]) else "SIN_DNI"
    nombre_pdf = f'{CARPETA_CERTIFICADOS}/certificado_{nombre}_{dni}.pdf'
    generar_pdf(row, nombre_pdf)

