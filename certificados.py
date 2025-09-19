import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib import colors
from datetime import datetime
from dateutil.relativedelta import relativedelta
import os, glob

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

def formato_fecha_actual():
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]
    hoy = datetime.today()
    return f"Lima, {hoy.day} de {meses[hoy.month-1]} del {hoy.year}"

def fecha_en_palabras(fecha_str):
    meses = [
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
    ]
    fecha = pd.to_datetime(fecha_str, dayfirst=True, errors="coerce")
    if pd.isna(fecha):
        return "fecha no especificada"
    return f"{fecha.day} de {meses[fecha.month-1]} del {fecha.year}"

def generar_pdf(data, nombre_archivo):
    doc = SimpleDocTemplate(nombre_archivo, pagesize=A4,
                            rightMargin=50, leftMargin=50,
                            topMargin=50, bottomMargin=50)
    elementos = []

    # --- Estilos ---
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Titulo", alignment=TA_CENTER, fontSize=14, spaceAfter=20, leading=16))
    styles.add(ParagraphStyle(name="Texto", alignment=TA_JUSTIFY, fontSize=11, leading=15))
    styles.add(ParagraphStyle(name="Firma", alignment=TA_CENTER, fontSize=11, spaceBefore=40))

    # --- Encabezado con logo ---
    logo_path = "logo_crea.png"
    if os.path.exists(logo_path):
        elementos.append(Image(logo_path, width=120, height=40))
    elementos.append(Spacer(1, 20))

    # --- Fecha ---
    elementos.append(Paragraph(formato_fecha_actual(), styles["Normal"]))
    elementos.append(Spacer(1, 20))

    # --- Título ---
    elementos.append(Paragraph("CERTIFICADO DE VOLUNTARIADO", styles["Titulo"]))

    # --- Texto principal con fechas en palabras ---
    fecha_vinculacion = fecha_en_palabras(data["Fecha de vinculación a Crea+ Perú:"])
    fecha_desvinculacion = fecha_en_palabras(data["Fecha de desvinculación a Crea+ Perú:"])
    tiempo_voluntariado = calcular_tiempo(
        data["Fecha de vinculación a Crea+ Perú:"],
        data["Fecha de desvinculación a Crea+ Perú:"]
    )

    texto = (
        "CREA MÁS PERU (en adelante, Crea+) es una asociación civil sin fines de lucro compuesta "
        "por un equipo multidisciplinario de jóvenes, el cual busca transformar la sociedad a través "
        "de una transformación personal de beneficiarios y voluntarios, otorgando herramientas para "
        "el crecimiento a través de un voluntariado profesional.<br/><br/>"
        f"Mediante el presente, Crea+ deja constancia que <b>{data['NOMBRE_SUNAT']}</b> con DNI <b>{data['DNI']}</b>, "
        f"participó como voluntaria/o desde el <b>{fecha_vinculacion}</b> "
        f"al <b>{fecha_desvinculacion}</b> en el rol de <b>{data['¿Qué rol desarrollaste dentro de la organización?']}</b>, "
        f"cumpliendo con {tiempo_voluntariado}.<br/><br/>"
        f"Certificamos que <b>{data['NOMBRE_SUNAT']}</b> demostró responsabilidad y compromiso en el desarrollo de sus funciones.<br/><br/>"
        "Se expide el presente certificado para los fines que se estimen convenientes.<br/><br/>"
        "Atentamente,"
    )

    elementos.append(Paragraph(texto, styles["Texto"]))
    elementos.append(Spacer(1, 30))

    # --- Firma y pie de página (sello) en una sola línea ---
    firma_path = "firma.png"
    pie_path = "pie_pagina.png"
    fila_imagenes = []
    if os.path.exists(firma_path):
        fila_imagenes.append(Image(firma_path, width=80, height=40))
    else:
        fila_imagenes.append(Spacer(1, 40))
    if os.path.exists(pie_path):
        fila_imagenes.append(Image(pie_path, width=120, height=40))
    else:
        fila_imagenes.append(Spacer(1, 40))
    tabla_firma = Table([fila_imagenes], colWidths=[200, 200])
    tabla_firma.setStyle(TableStyle([
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        # Si quieres borde, agrega: ("BOX", (0,0), (-1,-1), 0.5, colors.grey),
    ]))
    elementos.append(tabla_firma)
    elementos.append(Spacer(1, 10))

    # --- Nombre y cargo ---
    elementos.append(Paragraph("Diego Cabrera Zanatta<br/>Coordinador de Gestión de Talento Humano", styles["Firma"]))

    # --- Exportar PDF ---
    doc.build(elementos)

# =====================
# Bucle principal
# =====================
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
