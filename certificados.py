from reportlab.pdfgen import canvas 
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import getSampleStyleSheet
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

    # Fondo con la plantilla (usar JPG para evitar fondo negro)
    plantilla_path = "plantilla_certificado.jpg"
    if os.path.exists(plantilla_path):
        fondo = ImageReader(plantilla_path)
        c.drawImage(fondo, 0, 0, width=w, height=h)

    # Fecha actual
    c.setFont("Helvetica", 12)
    c.drawString(w-250, h-100, formato_fecha_actual())

    # Texto dinámico principal
    fecha_vinculacion = fecha_en_palabras(data["Fecha de vinculación a Crea+ Perú:"])
    fecha_desvinculacion = fecha_en_palabras(data["Fecha de desvinculación a Crea+ Perú:"])
    tiempo_voluntariado = calcular_tiempo(
        data["Fecha de vinculación a Crea+ Perú:"],
        data["Fecha de desvinculación a Crea+ Perú:"]
    )

    texto = (
        f"Mediante el presente, Crea+ deja constancia que <b>{data['NOMBRE_SUNAT']}</b> "
        f"con DNI <b>{data['DNI']}</b>, participó como voluntaria/o desde el <b>{fecha_vinculacion}</b> "
        f"al <b>{fecha_desvinculacion}</b> en el rol de <b>{data['¿Qué rol desarrollaste dentro de la organización?']}</b>, "
        f"cumpliendo con {tiempo_voluntariado}."
    )

    # Texto para reemplazar lo amarillo
    texto2 = (
        f"Certificamos que <b>{data['NOMBRE_SUNAT']}</b> demostró responsabilidad y "
        f"compromiso en el desarrollo de sus funciones."
    )

    # Estilos
    styles = getSampleStyleSheet()
    styleN = styles["Normal"]
    styleN.fontName = "Helvetica"
    styleN.fontSize = 12
    styleN.leading = 16

    # Párrafo 1
    P = Paragraph(texto, styleN)
    frame = Frame(70, h/2-80, w-140, 200, showBoundary=0)
    frame.addFromList([P], c)

    # Párrafo 2 (reemplazo de amarillo)
    P2 = Paragraph(texto2, styleN)
    frame2 = Frame(70, h/2-160, w-140, 100, showBoundary=0)
    frame2.addFromList([P2], c)

    # Firma
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(w/2, 120, "Diego Cabrera Zanatta")
    c.drawCentredString(w/2, 105, "Coordinador de Gestión de Talento Humano")

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
