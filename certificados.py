from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
import pandas as pd
import os, glob
from datetime import datetime
from dateutil.relativedelta import relativedelta

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

def generar_pdf(data, nombre_archivo):
    # Tamaño A4
    w, h = A4
    c = canvas.Canvas(nombre_archivo, pagesize=A4)

    # --- Fondo plantilla ---
    plantilla_path = "plantilla_certificado.png"
    if os.path.exists(plantilla_path):
        c.drawImage(plantilla_path, 0, 0, width=w, height=h)

    # --- Texto dinámico ---
    fecha_vinculacion = fecha_en_palabras(data["Fecha de vinculación a Crea+ Perú:"])
    fecha_desvinculacion = fecha_en_palabras(data["Fecha de desvinculación a Crea+ Perú:"])
    tiempo_voluntariado = calcular_tiempo(
        data["Fecha de vinculación a Crea+ Perú:"],
        data["Fecha de desvinculación a Crea+ Perú:"]
    )

    # Estilo de letra
    c.setFont("Helvetica", 12)

    # Fecha actual
    c.drawString(350, h-100, formato_fecha_actual())  # ajusta coordenadas

    # Texto central
    texto = (
        f"Mediante el presente, Crea+ deja constancia que {data['NOMBRE_SUNAT']} "
        f"con DNI {data['DNI']}, participó como voluntaria/o desde el {fecha_vinculacion} "
        f"al {fecha_desvinculacion} en el rol de {data['¿Qué rol desarrollaste dentro de la organización?']}, "
        f"cumpliendo con {tiempo_voluntariado}."
    )
    # Ajuste de párrafos
    from reportlab.platypus import Paragraph
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Frame

    styles = getSampleStyleSheet()
    styleN = styles["Normal"]
    styleN.fontName = "Helvetica"
    styleN.fontSize = 12
    styleN.leading = 16

    P = Paragraph(texto, styleN)
    frame = Frame(70, h/2-50, w-140, 200, showBoundary=0)  # caja de texto central
    frame.addFromList([P], c)

    # Firma (puede ir en la plantilla como imagen)
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(w/2, 120, "Diego Cabrera Zanatta")
    c.drawCentredString(w/2, 105, "Coordinador de Gestión de Talento Humano")

    c.save()

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
