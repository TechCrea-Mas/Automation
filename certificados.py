def formatear_nombre(texto):
    if pd.isna(texto):
        return ""
    # Convierte a título (ejemplo: "EVA MARIA SALOME" -> "Eva Maria Salome")
    return str(texto).title()

def generar_pdf(data, nombre_archivo):
    w, h = A4
    c = canvas.Canvas(nombre_archivo, pagesize=A4)
    
    # Fondo institucional
    plantilla_path = "plantilla_certificado.jpg"
    if os.path.exists(plantilla_path):
        fondo = ImageReader(plantilla_path)
        c.drawImage(fondo, 0, 0, width=w, height=h)

    # Estilos
    styles = getSampleStyleSheet()
    styleTitulo = ParagraphStyle(
        'titulo',
        parent=styles['Heading1'],
        alignment=1,
        fontSize=16,
        fontName="Helvetica-Bold",
        spaceAfter=18
    )
    styleNormal = ParagraphStyle(
        'normal',
        fontName="Helvetica",
        fontSize=12,
        leading=16,
        alignment=4,
        spaceAfter=12
    )

    # Variables formateadas
    nombre = formatear_nombre(data["NOMBRE_SUNAT"])
    area = formatear_nombre(data["¿En qué área o equipo participaste?"])
    rol = formatear_nombre(data["¿Qué rol desarrollaste dentro de la organización?"])
    dni = str(int(data["DNI"])) if not pd.isna(data["DNI"]) else ""

    fecha_vinculacion = fecha_en_palabras(data["Fecha de vinculación a Crea+ Perú:"])
    fecha_desvinculacion = fecha_en_palabras(data["Fecha de desvinculación a Crea+ Perú:"])
    tiempo_voluntariado = calcular_tiempo(
        data["Fecha de vinculación a Crea+ Perú:"],
        data["Fecha de desvinculación a Crea+ Perú:"]
    )

    # Fecha actual centrada arriba
    c.setFont("Helvetica", 12)
    c.drawCentredString(w/2, h-115, formato_fecha_actual())

    # Título
    P_titulo = Paragraph("CERTIFICADO DE VOLUNTARIADO", styleTitulo)
    frame_titulo = Frame(0, h-160, w, 30, showBoundary=0)
    frame_titulo.addFromList([P_titulo], c)

    # Texto certificado
    texto = (
        "<b>CREA MÁS PERU</b> (en adelante, Crea+) es una asociación civil sin fines de lucro compuesta por un equipo multidisciplinario de jóvenes, el cual busca transformar la sociedad a través de una transformación personal de beneficiarios y voluntarios, otorgando herramientas para el crecimiento a través de un voluntariado profesional.<br/><br/>"
        f"Mediante el presente, Crea+ deja constancia que <b>{nombre}</b> con DNI <b>{dni}</b>, participó como voluntaria/o en el área/equipo <b>{area}</b> desde el <b>{fecha_vinculacion}</b> al <b>{fecha_desvinculacion}</b> en el rol de <b>{rol}</b>, cumpliendo con {tiempo_voluntariado}.<br/><br/>"
        f"Certificamos que <b>{nombre}</b> demostró responsabilidad y compromiso en el desarrollo de sus funciones.<br/><br/>"
        "Se expide el presente certificado para los fines que se estimen convenientes.<br/><br/>"
        "Atentamente,"
    )

    frame = Frame(80, 220, w-160, 330, showBoundary=0)
    P = Paragraph(texto, styleNormal)
    frame.addFromList([P], c)

    c.save()
